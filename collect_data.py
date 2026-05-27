#!/usr/bin/env python3
"""
YouTube Firework Video Data Collector
======================================
Collects cake (组合烟花) and fountain (喷花) firework videos from YouTube
for online consumer preference research.
Uses yt-dlp for searching and metadata extraction — no YouTube API key required.

Usage:
    pip install -r requirements.txt
    python collect_data.py                          # direct connection
    python collect_data.py --proxy socks5://127.0.0.1:1080   # with SOCKS5 proxy
    python collect_data.py --proxy http://127.0.0.1:7890     # with HTTP proxy

Note for users in mainland China:
    YouTube is blocked. You MUST use a VPN/proxy and pass it via --proxy,
    or configure yt-dlp globally in ~/.config/yt-dlp/config
    (see https://github.com/yt-dlp/yt-dlp#configuration)

Output:
    firework_video_data.xlsx — formatted Excel with all collected data
"""

import os
import shutil
import subprocess
import json
import sys
import time
import re
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# ============================================================
# YT-DLP DISCOVERY
# ============================================================

def _find_ytdlp():
    """Find yt-dlp executable path, checking multiple locations."""
    # 1. Direct PATH lookup
    found = shutil.which("yt-dlp") or shutil.which("yt-dlp.exe")
    if found:
        return found

    # 2. Common install locations
    candidates = []
    for base in os.environ.get("PATH", "").split(os.pathsep):
        candidates.append(os.path.join(base, "yt-dlp.exe"))
        candidates.append(os.path.join(base, "yt-dlp"))
    # 3. Python Scripts directory (pip install --user or system)
    local_appdata = os.environ.get("LOCALAPPDATA", "")
    appdata = os.environ.get("APPDATA", "")
    for py_ver in ["Python312", "Python311", "Python310", "Python39", "Python3"]:
        for local_base in [
            os.path.join(local_appdata, "Programs", "Python", py_ver, "Scripts"),
            os.path.join(appdata, "Python", py_ver, "Scripts"),
            os.path.join("C:\\", "Program Files", py_ver, "Scripts"),
        ]:
            candidates.append(os.path.join(local_base, "yt-dlp.exe"))
            candidates.append(os.path.join(local_base, "yt-dlp"))

    for c in candidates:
        if os.path.isfile(c):
            return c

    return "yt-dlp"  # fallback, will fail with helpful error


YTDLP = _find_ytdlp()
PROXY = None  # set via --proxy CLI argument


# ============================================================
# CONFIGURATION
# ============================================================

OUTPUT_FILE = "firework_video_data.xlsx"
MAX_RESULTS_PER_QUERY = 25
FETCH_DELAY = 0.8  # seconds between detailed fetches to be respectful

# Search queries covering different angles
SEARCH_QUERIES = [
    "firework cake barrage repeater multi-shot demo test",
    "firework fountain ground fountain consumer review",
    "cake firework 500g 200g backyard test demo",
    "fountain firework cone ice fountain test demo",
    "multi-shot aerial cake firework review unboxing",
    "barrage repeater firework cake consumer demo",
    "consumer firework cake fountain unboxing test",
]

# Keywords suggesting the video IS likely a cake/fountain demo
INCLUDE_KEYWORDS = [
    "cake", "barrage", "repeater", "multi-shot", "multi shot",
    "fountain", "ground fountain", "aerial cake",
    "firework cake", "firework fountain", "500g cake", "200g cake",
    "consumer firework", "backyard firework", "cake demo",
    "cone fountain", "ice fountain", "compound cake",
]

# Keywords suggesting the video should be EXCLUDED
EXCLUDE_KEYWORDS = [
    "firework show", "fireworks show", "display shell",
    "professional display", "pyromusical", "firework competition",
    "firework festival", "music video", "how to make",
    "tutorial", "diy firework", "manufacturing", "drone",
    "wedding", "new year countdown", "new years eve",
    "4th of july show", "fourth of july show",
    "independence day show", "mall show", "stadium",
    "orchestra", "concert",
]

# ============================================================
# EXCEL HEADERS (strict order matching spec)
# ============================================================

HEADERS = [
    "样本编号",        # 1: Sample ID
    "视频URL",         # 2: Video URL
    "UP主名称",        # 3: Channel name
    "UP主类型",        # 4: Creator type (厂商/测评/个人) - MANUAL
    "发布年份",        # 5: Publish year
    "发布月份",        # 6: Publish month
    "产品大类",        # 7: Product (cake/fountain)
    "发数",            # 8: Shot count - MANUAL
    "筒径规格",        # 9: Tube diameter - MANUAL
    "药量(g)",         # 10: Powder weight (g) - MANUAL
    "燃放时长",        # 11: Duration category - MANUAL (≤10s/10-30s/30-60s/＞60s)
    "效果类型",        # 12: Effect type - MANUAL (爆响/单色/多色渐变/闪光/冷光/造型)
    "播放量",          # 13: View count
    "点击率CTR",       # 14: CTR - UNAVAILABLE (YouTube Studio only)
    "完播率",          # 15: Completion rate - UNAVAILABLE (YouTube Studio only)
    "点赞数",          # 16: Like count
    "评论数",          # 17: Comment count
    "收藏数",          # 18: Save count - UNAVAILABLE (not publicly exposed)
    "综合互动率",      # 19: Engagement rate = (likes+comments+saves)/views
]

MANUAL_COLS = {4, 8, 9, 10, 11, 12}  # 1-indexed columns needing manual annotation
UNAVAILABLE_COLS = {14, 15, 18}       # 1-indexed columns not publicly available
ENGAGEMENT_COL = 19                   # Formula column


# ============================================================
# YOUTUBE DATA COLLECTION
# ============================================================

def search_youtube(query, max_results=MAX_RESULTS_PER_QUERY):
    """Search YouTube via yt-dlp and return list of video info dicts."""
    videos = []
    cmd = [
        YTDLP,
        f"ytsearch{max_results}:{query}",
        "--dump-json",
        "--skip-download",
        "--no-warnings",
        "--ignore-errors",
        "--socket-timeout", "30","--extractor-args", "youtube:player_client=android,ios",
    ]
    if PROXY:
        cmd.extend(["--proxy", PROXY])
    try:
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
        if result.returncode != 0 and not result.stdout.strip():
            print(f"    Warning: search failed: {result.stderr[:150]}")
            return videos

        for line in result.stdout.strip().split("\n"):
            if not line:
                continue
            try:
                videos.append(json.loads(line))
            except json.JSONDecodeError:
                continue
    except FileNotFoundError:
        print(f"ERROR: yt-dlp not found at '{YTDLP}'. Run: pip install yt-dlp")
        sys.exit(1)
    except subprocess.TimeoutExpired:
        print(f"    Warning: search timed out")
    return videos


def get_detailed_info(video_id):
    """Fetch full video metadata from YouTube (view_count, like_count, etc.)."""
    time.sleep(FETCH_DELAY)
    cmd = [
        YTDLP,
        f"https://www.youtube.com/watch?v={video_id}",
        "--dump-json",
        "--skip-download",
        "--no-warnings",
        "--ignore-errors",
        "--socket-timeout", "30","--extractor-args", "youtube:player_client=android,ios",
    ]
    if PROXY:
        cmd.extend(["--proxy", PROXY])
    try:
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
        if result.returncode != 0 or not result.stdout.strip():
            return None
        return json.loads(result.stdout)
    except (subprocess.TimeoutExpired, json.JSONDecodeError):
        return None


def is_relevant(info):
    """Return True if the video is likely a real cake/fountain demo (not show/ad/compilation)."""
    title = (info.get("title") or "").lower()
    description = (info.get("description") or "").lower()
    text = f"{title} {description}"

    has_include = any(kw.lower() in text for kw in INCLUDE_KEYWORDS)
    if not has_include:
        return False

    has_exclude = any(kw.lower() in text for kw in EXCLUDE_KEYWORDS)
    if has_exclude:
        return False

    # Duration filter: real demos typically 30s–20min
    duration = info.get("duration") or 0
    if duration < 25 or duration > 1200:
        return False

    # Exclude Shorts
    if info.get("webpage_url", "").find("/shorts/") != -1:
        return False

    return True


def classify_product(title, description):
    """Classify as cake or fountain based on title/description keywords."""
    text = f"{title or ''} {description or ''}".lower()

    fountain_kw = ["fountain", "ground fountain", "cone fountain",
                   "ice fountain", "roman candle fountain"]
    cake_kw = ["cake", "barrage", "repeater", "multi-shot", "multi shot",
               "aerial cake", "500g cake", "200g cake", "compound cake",
               "9 shot", "12 shot", "16 shot", "25 shot", "36 shot",
               "49 shot", "100 shot", "144 shot"]

    for kw in fountain_kw:
        if kw in text:
            return "fountain"
    for kw in cake_kw:
        if kw in text:
            return "cake"
    return "未知"


def parse_date(upload_date_str):
    """Parse YYYYMMDD string into (year, month). Returns (None, None) on failure."""
    if not upload_date_str or len(upload_date_str) != 8:
        return None, None
    try:
        dt = datetime.strptime(upload_date_str, "%Y%m%d")
        return str(dt.year), f"{dt.month:02d}"
    except ValueError:
        return None, None


# ============================================================
# EXCEL OUTPUT
# ============================================================

def create_excel_template(filepath):
    """Create a formatted Excel workbook with headers and a notes sheet."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "烟花视频数据"

    header_font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )

    for col_idx, header in enumerate(HEADERS, 1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = border

    col_widths = {
        1: 10, 2: 45, 3: 28, 4: 12, 5: 10, 6: 10,
        7: 12, 8: 10, 9: 12, 10: 10, 11: 18, 12: 24,
        13: 14, 14: 12, 15: 10, 16: 12, 17: 12, 18: 12, 19: 18,
    }
    for col, width in col_widths.items():
        ws.column_dimensions[get_column_letter(col)].width = width

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(HEADERS))}1"

    # Notes sheet
    ws2 = wb.create_sheet("说明")
    ws2.column_dimensions["A"].width = 22
    ws2.column_dimensions["B"].width = 65

    label_font = Font(name="Arial", size=10, bold=True)
    value_font = Font(name="Arial", size=10)
    warn_font = Font(name="Arial", size=10, color="CC0000")

    notes = [
        ("数据采集日期", datetime.now().strftime("%Y-%m-%d %H:%M")),
        ("数据来源", "YouTube 公开视频"),
        ("采集方式", "yt-dlp 搜索 + 元数据提取（无需 YouTube API Key）"),
        ("时间范围", "2023–2026 年发布的视频"),
        ("产品类别", "cake（组合烟花 / 连发 / barrage / repeater）\nfountain（喷花 / 地面喷泉）"),
        ("", ""),
        ("═══ 字段说明 ═══", ""),
        ("样本编号", "系统自动编号"),
        ("视频URL", "YouTube 视频链接"),
        ("UP主名称", "频道名称"),
        ("UP主类型", "厂商 / 测评 / 个人 —— 需人工判断后填写"),
        ("发布年份 / 月份", "从视频元数据自动提取"),
        ("产品大类", "cake / fountain，基于标题关键词自动分类，建议人工复核"),
        ("发数", "⚠ 需人工观看视频后标注（如 36发、100发）"),
        ("筒径规格", "⚠ 需人工观看视频后标注（如 1英寸、1.5英寸）"),
        ("药量(g)", "⚠ 需人工观看视频后标注"),
        ("燃放时长", "⚠ 需人工标注：≤10s / 10-30s / 30-60s / ＞60s"),
        ("效果类型", "⚠ 需人工标注：爆响 / 单色 / 多色渐变 / 闪光 / 冷光 / 造型"),
        ("播放量", "YouTube 公开数据（随采集时间变化）"),
        ("点击率CTR", "❌ 仅视频上传者可在 YouTube Studio 查看，不可公开获取"),
        ("完播率", "❌ 仅视频上传者可在 YouTube Studio 查看，不可公开获取"),
        ("点赞数", "YouTube 公开数据（部分视频可能隐藏）"),
        ("评论数", "YouTube 公开数据（部分视频可能关闭评论）"),
        ("收藏数", "❌ YouTube 不公开收藏/保存数据"),
        ("综合互动率", "= (点赞数 + 评论数 + 收藏数) / 播放量"),
        ("", ""),
        ("═══ 重要提示 ═══", ""),
        ("数据局限性", "CTR、完播率、收藏数均不可公开获取，相应列留空。\n热度数据为采集时刻的快照，后续会变化。"),
        ("标注工作量", "发数/筒径/药量/燃放时长/效果类型需逐一观看视频后人工标注。"),
        ("代表性声明", "数据仅代表 YouTube 线上关注度偏好，不等同于实际购买行为。"),
        ("筛选说明", "已自动剔除 firework show / 广告混剪 / DIY教程 / Shorts 等无关视频。\n误剔除或遗漏的样本请手动补充。"),
    ]

    for row_idx, (label, value) in enumerate(notes, 1):
        c1 = ws2.cell(row=row_idx, column=1, value=label)
        c2 = ws2.cell(row=row_idx, column=2, value=value)
        if label.startswith("═══"):
            c1.font = Font(name="Arial", size=10, bold=True, color="2F5496")
        elif label.startswith("❌") or label.startswith("⚠"):
            c2.font = warn_font
        else:
            c1.font = label_font
        c2.font = value_font if not label.startswith("❌") else warn_font

    wb.save(filepath)
    print(f"  Excel template created: {filepath}")


def append_data(filepath, rows):
    """Append data rows to the Excel data sheet."""
    wb = openpyxl.load_workbook(filepath)
    ws = wb["烟花视频数据"]
    start_row = ws.max_row + 1

    data_font = Font(name="Arial", size=10)
    data_align = Alignment(vertical="center")
    url_font = Font(name="Arial", size=10, color="0563C1", underline="single")
    border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    manual_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    unavailable_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

    for row_idx, row_data in enumerate(rows, start_row):
        for col_idx, value in enumerate(row_data, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.font = url_font if col_idx == 2 else data_font
            cell.alignment = data_align
            cell.border = border

            if col_idx in MANUAL_COLS:
                cell.fill = manual_fill
            elif col_idx in UNAVAILABLE_COLS:
                cell.fill = unavailable_fill

    wb.save(filepath)
    print(f"  {len(rows)} data rows appended (starting row {start_row})")


# ============================================================
# MAIN
# ============================================================

def collect_data(output_file):
    """Run the full collection pipeline."""
    print("=" * 60)
    print("  YouTube Firework Video Data Collector")
    print("  cake (组合烟花) & fountain (喷花) | 2023–2026")
    print("=" * 60)

    # ---- Phase 1: Search ----
    print("\n[Phase 1] Searching YouTube...")
    all_videos = {}

    for query in SEARCH_QUERIES:
        print(f"  Query: '{query[:60]}...'")
        results = search_youtube(query)
        print(f"    -> {len(results)} results")
        for info in results:
            vid = info.get("id")
            if vid and vid not in all_videos:
                all_videos[vid] = info

    print(f"\n  Unique videos (before filtering): {len(all_videos)}")

    # ---- Phase 2: Filter ----
    print("\n[Phase 2] Filtering...")
    filtered = {vid: info for vid, info in all_videos.items() if is_relevant(info)}
    print(f"  After filtering: {len(filtered)} videos")

    # ---- Phase 3: Enrich with detailed metadata ----
    print("\n[Phase 3] Fetching detailed metadata (views, likes, comments)...")
    print(f"  (Delay: {FETCH_DELAY}s per video — this may take a few minutes)\n")

    rows = []
    sample_id = 0
    skipped_year = 0
    fetch_fail = 0

    for idx, (video_id, info) in enumerate(filtered.items()):
        title_short = (info.get("title") or "N/A")[:55]
        print(f"  [{idx+1}/{len(filtered)}] {title_short}...", end=" ", flush=True)

        detailed = get_detailed_info(video_id)
        if detailed:
            info = detailed
        else:
            fetch_fail += 1
            print("(skipped — fetch failed)")
            continue

        # Date filter: 2023–2026
        year, month = parse_date(info.get("upload_date"))
        if year is None:
            print("(skipped — bad date)")
            continue
        if int(year) < 2023 or int(year) > 2026:
            skipped_year += 1
            print(f"(skipped — year={year})")
            continue

        title = info.get("title") or "N/A"
        description = info.get("description") or ""
        channel = info.get("uploader") or info.get("channel") or "N/A"
        view_count = info.get("view_count") or 0
        like_count = info.get("like_count") or 0
        comment_count = info.get("comment_count") or 0
        product_type = classify_product(title, description)
        engagement = round((like_count + comment_count) / view_count, 8) if view_count > 0 else 0

        sample_id += 1
        rows.append([
            sample_id,
            f"https://www.youtube.com/watch?v={video_id}",
            channel,
            "",              # UP主类型 — manual
            year,
            month,
            product_type,
            "",              # 发数 — manual
            "",              # 筒径规格 — manual
            "",              # 药量(g) — manual
            "",              # 燃放时长 — manual
            "",              # 效果类型 — manual
            view_count,
            "",              # CTR — unavailable
            "",              # 完播率 — unavailable
            like_count,
            comment_count,
            "",              # 收藏数 — unavailable
            engagement,
        ])
        print("ok")

    # ---- Phase 4: Write Excel ----
    print(f"\n[Phase 4] Writing Excel...")
    print(f"  Valid samples: {sample_id}")
    print(f"  Skipped (year out of range): {skipped_year}")
    print(f"  Skipped (fetch failed): {fetch_fail}")

    if not Path(output_file).exists():
        create_excel_template(output_file)
    else:
        print(f"  Using existing template: {output_file}")

    if rows:
        append_data(output_file, rows)

    # ---- Summary ----
    cakes = sum(1 for r in rows if r[6] == "cake")
    fountains = sum(1 for r in rows if r[6] == "fountain")
    unknown = sum(1 for r in rows if r[6] == "未知")

    print("\n" + "=" * 60)
    print("  COLLECTION COMPLETE")
    print(f"  Output file:  {output_file}")
    print(f"  Total rows:   {len(rows)}")
    print(f"    Cake:       {cakes}")
    print(f"    Fountain:   {fountains}")
    print(f"    Unknown:    {unknown}")
    print(f"\n  Columns needing manual annotation (yellow highlight):")
    print(f"    UP主类型 / 发数 / 筒径规格 / 药量(g) / 燃放时长 / 效果类型")
    print(f"  Columns NOT publicly available (grey highlight):")
    print(f"    点击率CTR / 完播率 / 收藏数")
    print("=" * 60)


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(
        description="YouTube Firework Video Data Collector (cake & fountain)"
    )
    parser.add_argument(
        "--proxy", type=str, default=None,
        help="Proxy URL for yt-dlp, e.g. socks5://127.0.0.1:1080 or http://127.0.0.1:7890"
    )
    args = parser.parse_args()
    PROXY = args.proxy

    output_path = str(Path(__file__).parent / OUTPUT_FILE)
    if PROXY:
        print(f"Using proxy: {PROXY}")
    collect_data(output_path)
