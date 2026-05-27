#!/usr/bin/env python3
"""
YouTube Firework Video Data Collector
======================================
Collects cake (组合烟花) and fountain (喷花) firework videos from YouTube
for online consumer preference research.

Search: Invidious API (public, no auth required)
Enrichment: yt-dlp (fallback to Invidious data if blocked)

Usage:
    pip install -r requirements.txt
    python collect_data.py

Output:
    firework_video_data.xlsx — formatted Excel with all collected data
"""

import os
import shutil
import subprocess
import json
import sys
import time
from datetime import datetime
from pathlib import Path
from urllib.request import Request, urlopen
from urllib.error import URLError
from urllib.parse import quote

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# ============================================================
# YT-DLP DISCOVERY (for enrichment only)
# ============================================================

def _find_ytdlp():
    found = shutil.which("yt-dlp") or shutil.which("yt-dlp.exe")
    if found:
        return found
    local_appdata = os.environ.get("LOCALAPPDATA", "")
    appdata = os.environ.get("APPDATA", "")
    for py_ver in ["Python312", "Python311", "Python310", "Python39", "Python3"]:
        for base in [
            os.path.join(local_appdata, "Programs", "Python", py_ver, "Scripts"),
            os.path.join(appdata, "Python", py_ver, "Scripts"),
            os.path.join("C:\\", "Program Files", py_ver, "Scripts"),
        ]:
            for name in ["yt-dlp.exe", "yt-dlp"]:
                p = os.path.join(base, name)
                if os.path.isfile(p):
                    return p
    return "yt-dlp"

YTDLP = _find_ytdlp()


# ============================================================
# INVIDIOUS INSTANCES (public YouTube proxies — no auth needed)
# ============================================================

INVIDIOUS_INSTANCES = [
    "https://invidious.fdn.fr",
    "https://vid.puffyan.us",
    "https://invidious.perennialte.ch",
    "https://yewtu.be",
    "https://inv.nadeko.net",
    "https://invidious.privacyredirect.com",
    "https://invidious.nerdvpn.de",
    "https://invidious.slipfox.xyz",
]


def _invidious_search(query, max_results=25):
    """Search YouTube via Invidious API. Returns list of video dicts."""
    encoded = quote(query)
    for instance in INVIDIOUS_INSTANCES:
        url = f"{instance}/api/v1/search?q={encoded}&type=video&page=1"
        try:
            req = Request(url, headers={"User-Agent": "Mozilla/5.0"})
            resp = urlopen(req, timeout=20)
            data = json.loads(resp.read().decode())
            results = []
            for item in data[:max_results]:
                if item.get("type") != "video":
                    continue
                results.append({
                    "id": item.get("videoId"),
                    "title": item.get("title", ""),
                    "description": item.get("description", ""),
                    "uploader": item.get("author", ""),
                    "uploader_id": item.get("authorId", ""),
                    "duration": item.get("lengthSeconds", 0),
                    "view_count": item.get("viewCount", 0),
                    "upload_date": _iso_to_yyyymmdd(item.get("publishedText", "")),
                    "webpage_url": f"https://www.youtube.com/watch?v={item.get('videoId')}",
                })
            if results:
                return results
        except (URLError, json.JSONDecodeError, Exception) as e:
            continue
    return []


def _iso_to_yyyymmdd(published_text):
    """Convert Invidious published text (e.g. '2 years ago') to YYYYMMDD string."""
    # Invidious returns relative text. We use the video extraction date instead.
    return None


# ============================================================
# CONFIGURATION
# ============================================================

OUTPUT_FILE = "firework_video_data.xlsx"
MAX_RESULTS_PER_QUERY = 25
FETCH_DELAY = 0.8

SEARCH_QUERIES = [
    "firework cake barrage repeater multi-shot demo test",
    "firework fountain ground fountain consumer review",
    "cake firework 500g 200g backyard test demo",
    "fountain firework cone ice fountain test demo",
    "multi-shot aerial cake firework review unboxing",
    "barrage repeater firework cake consumer demo",
    "consumer firework cake fountain unboxing test",
]

INCLUDE_KEYWORDS = [
    "cake", "barrage", "repeater", "multi-shot", "multi shot",
    "fountain", "ground fountain", "aerial cake",
    "firework cake", "firework fountain", "500g cake", "200g cake",
    "consumer firework", "backyard firework", "cake demo",
    "cone fountain", "ice fountain", "compound cake",
]

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
# EXCEL HEADERS (strict order)
# ============================================================

HEADERS = [
    "样本编号", "视频URL", "UP主名称", "UP主类型",
    "发布年份", "发布月份", "产品大类",
    "发数", "筒径规格", "药量(g)", "燃放时长", "效果类型",
    "播放量", "点击率CTR", "完播率",
    "点赞数", "评论数", "收藏数", "综合互动率",
]

MANUAL_COLS = {4, 8, 9, 10, 11, 12}
UNAVAILABLE_COLS = {14, 15, 18}


# ============================================================
# ENRICHMENT (via yt-dlp)
# ============================================================

def get_detailed_info(video_id):
    """Fetch full video metadata via yt-dlp. Returns None on failure."""
    time.sleep(FETCH_DELAY)
    cmd = [
        YTDLP,
        f"https://www.youtube.com/watch?v={video_id}",
        "--dump-json", "--skip-download",
        "--no-warnings", "--ignore-errors",
        "--socket-timeout", "20",
        "--retries", "2",
        "--extractor-args", "youtube:player_client=android,ios",
    ]
    try:
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
        if result.returncode != 0 or not result.stdout.strip():
            return None
        return json.loads(result.stdout)
    except (subprocess.TimeoutExpired, json.JSONDecodeError):
        return None


# ============================================================
# FILTERING & CLASSIFICATION
# ============================================================

def is_relevant(info):
    title = (info.get("title") or "").lower()
    description = (info.get("description") or "").lower()
    text = f"{title} {description}"

    if not any(kw.lower() in text for kw in INCLUDE_KEYWORDS):
        return False
    if any(kw.lower() in text for kw in EXCLUDE_KEYWORDS):
        return False

    duration = info.get("duration") or 0
    if duration < 25 or duration > 1200:
        return False

    url = info.get("webpage_url", "")
    if "/shorts/" in url:
        return False

    return True


def classify_product(title, description):
    text = f"{title or ''} {description or ''}".lower()
    for kw in ["fountain", "ground fountain", "cone fountain", "ice fountain"]:
        if kw in text:
            return "fountain"
    for kw in ["cake", "barrage", "repeater", "multi-shot", "multi shot",
               "aerial cake", "500g cake", "200g cake", "compound cake"]:
        if kw in text:
            return "cake"
    return "未知"


def parse_date(upload_date_str):
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

    notes = [
        ("数据采集日期", datetime.now().strftime("%Y-%m-%d %H:%M")),
        ("数据来源", "YouTube 公开视频（通过 Invidious 镜像搜索）"),
        ("采集方式", "Invidious API 搜索 + yt-dlp 元数据增强"),
        ("时间范围", "2023–2026 年发布的视频"),
        ("产品类别", "cake（组合烟花 / 连发）\nfountain（喷花 / 地面喷泉）"),
        ("", ""),
        ("═══ 字段说明 ═══", ""),
        ("样本编号", "系统自动编号"),
        ("视频URL", "YouTube 视频链接"),
        ("UP主名称", "频道名称"),
        ("UP主类型", "厂商 / 测评 / 个人 —— 需人工判断后填写"),
        ("发布年份 / 月份", "从视频元数据自动提取"),
        ("产品大类", "cake / fountain，基于标题关键词自动分类，建议人工复核"),
        ("发数", "⚠ 需人工观看视频后标注"),
        ("筒径规格", "⚠ 需人工观看视频后标注"),
        ("药量(g)", "⚠ 需人工观看视频后标注"),
        ("燃放时长", "⚠ 需人工标注：≤10s / 10-30s / 30-60s / ＞60s"),
        ("效果类型", "⚠ 需人工标注：爆响 / 单色 / 多色渐变 / 闪光 / 冷光 / 造型"),
        ("播放量", "YouTube 公开数据"),
        ("点击率CTR", "❌ 仅视频上传者可在 YouTube Studio 查看"),
        ("完播率", "❌ 仅视频上传者可在 YouTube Studio 查看"),
        ("点赞数", "YouTube 公开数据"),
        ("评论数", "YouTube 公开数据"),
        ("收藏数", "❌ YouTube 不公开收藏/保存数据"),
        ("综合互动率", "= (点赞数 + 评论数 + 收藏数) / 播放量"),
        ("", ""),
        ("═══ 重要提示 ═══", ""),
        ("数据局限性", "CTR、完播率、收藏数均不可公开获取。热度数据为采集时刻的快照。"),
        ("标注工作量", "发数/筒径/药量/燃放时长/效果类型需逐一观看视频后人工标注。"),
        ("代表性声明", "数据仅代表 YouTube 线上关注度偏好，不等同于实际购买行为。"),
    ]

    for row_idx, (label, value) in enumerate(notes, 1):
        c1 = ws2.cell(row=row_idx, column=1, value=label)
        c2 = ws2.cell(row=row_idx, column=2, value=value)
        if label.startswith("═══"):
            c1.font = Font(name="Arial", size=10, bold=True, color="2F5496")
        c1.font = Font(name="Arial", size=10, bold=True)

    wb.save(filepath)
    print(f"  Excel template created: {filepath}")


def append_data(filepath, rows):
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
    print("=" * 60)
    print("  YouTube Firework Video Data Collector")
    print("  cake (组合烟花) & fountain (喷花) | 2023–2026")
    print("  Search: Invidious API | Enrichment: yt-dlp")
    print("=" * 60)

    if not Path(output_file).exists():
        create_excel_template(output_file)
    else:
        print(f"  Using existing template: {output_file}")

    # ---- Phase 1: Search via Invidious ----
    print("\n[Phase 1] Searching YouTube via Invidious API...")
    all_videos = {}

    for query in SEARCH_QUERIES:
        print(f"  Query: '{query[:60]}...'")
        results = _invidious_search(query, max_results=MAX_RESULTS_PER_QUERY)
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

    # ---- Phase 3: Enrich with yt-dlp (views, likes, comments, date) ----
    print("\n[Phase 3] Enriching metadata via yt-dlp (views, likes, dates)...")
    print(f"  (Delay: {FETCH_DELAY}s per video)\n")

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
            # yt-dlp failed — use Invidious data directly if available
            fetch_fail += 1
            print("(yt-dlp failed, using Invidious data)", end=" ")
            # Continue with the data we have from Invidious
            if not info.get("upload_date"):
                continue

        # Date filter: 2023–2026
        upload_date_str = info.get("upload_date")
        if not upload_date_str:
            print("(skipped — no date)")
            continue

        year, month = parse_date(upload_date_str)
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
            "",
            year,
            month,
            product_type,
            "", "", "", "", "",
            view_count,
            "", "",
            like_count,
            comment_count,
            "",
            engagement,
        ])
        print("ok")

    # ---- Phase 4: Write Excel ----
    print(f"\n[Phase 4] Writing Excel...")
    print(f"  Valid samples: {sample_id}")
    print(f"  Skipped (year out of range): {skipped_year}")
    print(f"  Skipped (fetch/enrich failed): {fetch_fail}")

    if rows:
        append_data(output_file, rows)
    else:
        print("  No data rows — check Invidious instance availability.")

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
    print(f"\n  Search method: Invidious API (public mirror, no auth)")
    print(f"  Enrich method: yt-dlp (with Android/iOS client fallback)")
    print("=" * 60)


if __name__ == "__main__":
    output_path = str(Path(__file__).parent / OUTPUT_FILE)
    collect_data(output_path)
