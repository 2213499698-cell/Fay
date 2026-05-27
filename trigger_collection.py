#!/usr/bin/env python3
"""
Trigger the YouTube firework data collection workflow on GitHub Actions,
wait for completion, and download the Excel result.

Requires GH_PAT environment variable (GitHub Personal Access Token with repo scope).
Create one at: https://github.com/settings/tokens/new?scopes=repo

Usage:
    GH_PAT=ghp_xxx python trigger_collection.py
"""

import os, sys, time, json, tempfile
from pathlib import Path
from urllib.request import Request, urlopen
from urllib.error import URLError
from datetime import datetime, timezone

REPO = "2213499698-cell/Fay"
WF_FILE = "collect.yml"
API_BASE = "https://api.github.com"

PAT = os.environ.get("GH_PAT") or os.environ.get("GITHUB_TOKEN")
if not PAT:
    print("ERROR: Set GH_PAT environment variable.")
    print("Create a token at: https://github.com/settings/tokens/new?scopes=repo")
    sys.exit(1)

HEADERS = {
    "Authorization": f"Bearer {PAT}",
    "Accept": "application/vnd.github+json",
    "User-Agent": "FireworkCollector",
}


def api(method, path, data=None):
    """Call GitHub REST API."""
    url = f"{API_BASE}{path}"
    body = json.dumps(data).encode() if data else None
    try:
        req = Request(url, data=body, headers=HEADERS, method=method)
        resp = urlopen(req, timeout=30)
        if resp.status == 204:
            return None
        return json.loads(resp.read().decode())
    except URLError as e:
        # GitHub API returns error body for 4xx/5xx
        if hasattr(e, "read"):
            print(f"  API error: {e.read().decode()[:300]}")
        else:
            print(f"  API error: {e}")
        return None


def main():
    print("=" * 55)
    print("  Triggering firework data collection...")
    print(f"  Repo: {REPO}")
    print("=" * 55)

    # Step 1: Trigger workflow via dispatches API
    print("\n[1] Starting workflow...")
    result = api("POST", f"/repos/{REPO}/actions/workflows/{WF_FILE}/dispatches",
                 {"ref": "master"})
    if result is None:
        # 204 No Content = success
        print("  Workflow triggered successfully.")
    else:
        print("  Failed to trigger workflow.")
        sys.exit(1)

    # Step 2: Wait a moment, then find the run
    print("\n[2] Waiting for run to start...")
    time.sleep(5)

    run_id = None
    for attempt in range(12):
        runs = api("GET", f"/repos/{REPO}/actions/runs?per_page=5&event=workflow_dispatch")
        if runs and "workflow_runs" in runs:
            for run in runs["workflow_runs"]:
                if run.get("status") in ("queued", "in_progress", "completed"):
                    run_id = run["id"]
                    break
        if run_id:
            break
        print("  Waiting for run to appear...")
        time.sleep(5)

    if not run_id:
        print("  ERROR: Could not find workflow run.")
        sys.exit(1)

    print(f"  Run ID: {run_id}")
    print(f"  URL: https://github.com/{REPO}/actions/runs/{run_id}")

    # Step 3: Poll until complete
    print("\n[3] Waiting for completion...")
    status = "in_progress"
    while status in ("queued", "in_progress"):
        time.sleep(15)
        run_info = api("GET", f"/repos/{REPO}/actions/runs/{run_id}")
        if not run_info:
            print("  Failed to fetch run status.")
            sys.exit(1)
        status = run_info.get("status", "unknown")
        conclusion = run_info.get("conclusion", "")
        print(f"  Status: {status} ({conclusion or '...'})")

    if conclusion != "success":
        print(f"\n  Workflow failed (conclusion: {conclusion}).")
        print(f"  Check logs: https://github.com/{REPO}/actions/runs/{run_id}")
        sys.exit(1)

    # Step 4: Download artifact
    print("\n[4] Downloading Excel file...")
    artifacts = api("GET", f"/repos/{REPO}/actions/runs/{run_id}/artifacts")
    if not artifacts or "artifacts" not in artifacts:
        print("  No artifacts found.")
        sys.exit(1)

    excel_artifact = None
    for a in artifacts["artifacts"]:
        if a["name"] == "firework_video_data":
            excel_artifact = a
            break

    if not excel_artifact:
        print("  firework_video_data artifact not found.")
        sys.exit(1)

    download_url = excel_artifact["archive_download_url"]
    print(f"  Artifact: {excel_artifact['name']} ({excel_artifact['size_in_bytes']} bytes)")

    # Download zip
    req = Request(download_url, headers=HEADERS)
    resp = urlopen(req, timeout=60)
    zip_data = resp.read()

    # Save to project directory
    output_dir = Path(__file__).parent
    timestamp = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    output_file = output_dir / f"firework_video_data_{timestamp}.xlsx"

    # Unzip and save
    import zipfile
    with tempfile.NamedTemporaryFile(suffix=".zip", delete=False) as tmp:
        tmp.write(zip_data)
        tmp_path = tmp.name

    with zipfile.ZipFile(tmp_path, "r") as zf:
        for name in zf.namelist():
            if name.endswith(".xlsx"):
                with zf.open(name) as src, open(output_file, "wb") as dst:
                    dst.write(src.read())
                break

    os.unlink(tmp_path)

    print(f"\n{'=' * 55}")
    print(f"  DONE!")
    print(f"  Saved: {output_file}")
    print(f"  Rows: open {output_file} to view")
    print(f"{'=' * 55}")


if __name__ == "__main__":
    main()
