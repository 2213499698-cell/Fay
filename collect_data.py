#!/usr/bin/env python3
"""
YouTube Firework Video Data Collector
======================================
Collects cake and fountain firework videos from YouTube (2023-2026)
using the official YouTube Data API v3.

Requires YOUTUBE_API_KEY environment variable (GitHub Secret).

Usage:
    YOUTUBE_API_KEY=xxx python collect_data.py

Output:
    firework_video_data.xlsx
"""

import os
import json
import sys
import time
from datetime import datetime
from pathlib import Path
from urllib.request import Request, urlopen
from urllib.error import URLError
from urllib.parse import quote, urlencode

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# ============================================================
# CONFIGURATION
# ============================================================

OUTPUT_FILE = "firework_video_data.xlsx"
MAX_RESULTS_PER_QUERY = 25

API_KEY = os.environ.get("YOUTUBE_API_KEY", "")

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
# YOUTUBE DATA API v3
# ============================================================

def _api_call(endpoint, params):
    """Call YouTube Data API v3. Returns parsed JSON or None."""
    params["key"] = API_KEY
    url = f"https://www.googleapis.com/youtube/v3/{endpoint}?{urlencode(params)}"
    try:
        req = Request(url, headers={"User-Agent": "FireworkDataCollector/1.0"})
        resp = urlopen(req, timeout=30)
        return json.loads(resp.read().decode())
    except URLError as e:
        print(f"    API error: {e}")
        return None


def search_youtube_api(query, max_results=25):
    """Search YouTube via API. Returns list of video IDs and snippet info."""
    results = []
    page_token = None

    while len(results) < max_results:
        params = {
            "part": "snippet",
            "q": query,
            "type": "video",
            "maxResults": min(50, max_results - len(results)),
            "relevanceLanguage": "en",
            "regionCode": "US",
        }
        if page_token:
            params["pageToken"] = page_token

        data = _api_call("search", params)
        if not data or "items" not in data:
            break

        for item in data["items"]:
            vid = item["id"].get("videoId")
            if not vid:
                continue
            snippet = item.get("snippet", {})
            results.append({
                "id": vid,
                "title": snippet.get("title", ""),
                "description": snippet.get("description", ""),
                "uploader": snippet.get("channelTitle", ""),
                "published_at": snippet.get("publishedAt", ""),
                "webpage_url": f"https://www.youtube.com/watch?v={vid}",
            })

        page_token = data.get("nextPageToken")
        if not page_token:
            break

    return results[:max_results]


def get_batch_details(video_ids):
    """Fetch statistics and contentDetails for up to 50 videos at once."""
    params = {
        "part": "statistics,contentDetails,snippet",
        "id": ",".join(video_ids),
        "maxResults": "50",
    }
    data = _api_call("videos", params)
    if not data or "items" not in data:
        return {}

    details = {}
    for item in data["items"]:
        vid = item["id"]
        snippet = item.get("snippet", {})
        stats = item.get("statistics", {})
        content = item.get("contentDetails", {})
        details[vid] = {
            "title": snippet.get("title", ""),
            "description": snippet.get("description", ""),
            "uploader": snippet.get("channelTitle", ""),
            "published_at": snippet.get("publishedAt", ""),
            "view_count": int(stats.get("viewCount", 0)),
            "like_count": int(stats.get("likeCount", 0)),
            "comment_count": int(stats.get("commentCount", 0)),
            "duration_iso": content.get("duration", ""),  # ISO 8601: PT3M45S
        }
    return details


def parse_iso_duration(iso_dur):
    """Convert ISO 8601 duration (PT3M45S) to seconds."""
    import re
    match = re.match(r'PT(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?', iso_dur)
    if not match:
        return 0
    h, m, s = match.groups()
    return int(h or 0) * 3600 + int(m or 0) * 60 + int(s or 0)


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

    duration = info.get("duration_seconds") or 0
    if duration and (duration < 25 or duration > 1200):
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


def parse_published(published_at):
    """Parse ISO 8601 datetime to (year, month)."""
    if not published_at:
        return None, None
    try:
        # 2023-01-15T12:30:00Z
        dt = datetime.strptime(published_at[:19], "%Y-%m-%dT%H:%M:%S")
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
    if not API_KEY:
        print("ERROR: YOUTUBE_API_KEY environment variable not set.")
        print("Set it via: YOUTUBE_API_KEY=xxx python collect_data.py")
        print("Or configure it as a GitHub Secret.")
        sys.exit(1)

    print("=" * 60)
    print("  YouTube Firework Video Data Collector")
    print("  YouTube Data API v3 | cake & fountain | 2023-2026")
    print("=" * 60)

    if not Path(output_file).exists():
        create_excel_template(output_file)
    else:
        # Always create fresh template to avoid dupes
        create_excel_template(output_file)

    # Phase 1: Search
    print("\n[Phase 1] Searching via YouTube Data API v3...")
    all_videos = {}
    total_quota = 0

    for query in SEARCH_QUERIES:
        print(f"  Query: '{query[:60]}...'")
        results = search_youtube_api(query, max_results=MAX_RESULTS_PER_QUERY)
        total_quota += 100  # search.list costs 100 units
        print(f"    -> {len(results)} results")
        for info in results:
            vid = info.get("id")
            if vid and vid not in all_videos:
                all_videos[vid] = info

    print(f"\n  Unique videos: {len(all_videos)}")
    print(f"  API quota used so far: {total_quota} / 10,000")

    # Phase 2: Filter by title/description keywords
    print("\n[Phase 2] Filtering...")
    filtered = {vid: info for vid, info in all_videos.items() if is_relevant(info)}
    print(f"  After filtering: {len(filtered)} videos")

    # Phase 3: Batch enrichment via videos.list
    print("\n[Phase 3] Fetching statistics via API (batch)...")
    rows = []
    sample_id = 0
    video_ids = list(filtered.keys())

    # Process in batches of 50
    for i in range(0, len(video_ids), 50):
        batch = video_ids[i:i + 50]
        details = get_batch_details(batch)
        total_quota += 1  # videos.list costs 1 unit per call

        for vid in batch:
            info = filtered[vid]
            detail = details.get(vid, {})

            if detail:
                info.update(detail)

            # Parse duration
            iso_dur = info.get("duration_iso", "")
            info["duration_seconds"] = parse_iso_duration(iso_dur) if iso_dur else 0

            # Re-filter with actual duration now available
            if not is_relevant(info):
                continue

            # Date filter
            year, month = parse_published(info.get("published_at", ""))
            if year is None:
                continue
            if int(year) < 2023 or int(year) > 2026:
                continue

            title = info.get("title") or "N/A"
            description = info.get("description") or ""
            channel = info.get("uploader") or "N/A"
            view_count = info.get("view_count") or 0
            like_count = info.get("like_count") or 0
            comment_count = info.get("comment_count") or 0
            product_type = classify_product(title, description)
            engagement = round((like_count + comment_count) / view_count, 8) if view_count > 0 else 0

            sample_id += 1
            rows.append([
                sample_id,
                f"https://www.youtube.com/watch?v={vid}",
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

    print(f"  Total API quota used: {total_quota} / 10,000")

    # Phase 4: Write Excel
    print(f"\n[Phase 4] Writing Excel...")
    print(f"  Valid samples: {len(rows)}")

    if rows:
        append_data(output_file, rows)
    else:
        print("  No data rows. Check API key or quota.")

    cakes = sum(1 for r in rows if r[6] == "cake")
    fountains = sum(1 for r in rows if r[6] == "fountain")

    print("\n" + "=" * 60)
    print(f"  DONE - {len(rows)} samples")
    print(f"    Cake: {cakes}  Fountain: {fountains}")
    print(f"    API quota used: {total_quota}/10000")
    print(f"    Output: {output_file}")
    print("=" * 60)


if __name__ == "__main__":
    output_path = str(Path(__file__).parent / OUTPUT_FILE)
    collect_data(output_path)
