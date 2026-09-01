#!/usr/bin/env python3
"""
CES Weekly Update -> BigQuery Sync
====================================
Flattens history.json (one row per card) and loads it into BigQuery,
so the whole update history is queryable with SQL.

Table:   wmt-d2-prod.ces_weekly_updates_v0c003n.updates
Mode:    Full replace each run (history.json is the single source of truth;
         this script never invents data, it mirrors it).

Usage:
    python3 sync_bigquery.py           # sync now
    python3 sync_bigquery.py --dry     # print rows, don't load
"""

import argparse
import json
import subprocess
import sys
import tempfile
from datetime import datetime, timezone
from pathlib import Path

HERE = Path(__file__).parent
HISTORY_F = HERE / "history.json"

PROJECT = "wmt-d2-prod"
DATASET = "ces_weekly_updates_v0c003n"
TABLE = "updates"


def week_start_date(label: str) -> str:
    """Parse 'Week of July 25, 2026' -> '2026-07-25'. Returns '' if unparseable."""
    try:
        dt = datetime.strptime(label, "Week of %B %d, %Y")
        return dt.strftime("%Y-%m-%d")
    except ValueError:
        return ""


def flatten(history: dict) -> list[dict]:
    """One row per card, across all weeks."""
    now = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S")
    rows = []
    for week_key, wk in history.items():
        wsd = week_start_date(wk["label"])
        for i, card in enumerate(wk["cards"]):
            rows.append({
                "week_key": week_key,
                "week_label": wk["label"],
                "week_start_date": wsd,
                "section": card.get("section", ""),
                "title": card.get("title", ""),
                "body_text": "\n\n".join(card.get("text", [])),
                "next_steps": card.get("ns", ""),
                "report_url": card.get("report", ""),
                "card_index": i,
                "loaded_at": now,
            })
    return rows


def load_to_bigquery(rows: list[dict]) -> None:
    with tempfile.NamedTemporaryFile(mode="w", suffix=".jsonl", delete=False) as f:
        for row in rows:
            f.write(json.dumps(row, ensure_ascii=False) + "\n")
        ndjson_path = f.name

    cmd = [
        "bq", f"--project_id={PROJECT}", "load",
        "--source_format=NEWLINE_DELIMITED_JSON",
        "--replace",
        f"{DATASET}.{TABLE}",
        ndjson_path,
    ]
    print(f"Loading {len(rows)} rows into {PROJECT}:{DATASET}.{TABLE} ...")
    result = subprocess.run(cmd, capture_output=True, text=True)
    print(result.stdout)
    if result.returncode != 0:
        sys.exit(f"BigQuery load FAILED:\n{result.stderr}")
    print("Sync complete.")


def main():
    parser = argparse.ArgumentParser(description="Sync ces-weekly-update history.json to BigQuery")
    parser.add_argument("--dry", action="store_true", help="Print rows without loading to BigQuery")
    args = parser.parse_args()

    if not HISTORY_F.exists():
        sys.exit(f"ERROR: {HISTORY_F} not found")

    history = json.loads(HISTORY_F.read_text(encoding="utf-8"))
    rows = flatten(history)
    if not rows:
        sys.exit("ERROR: No rows to sync (history.json is empty)")

    print(f"Flattened {len(rows)} cards across {len(history)} weeks:")
    for wk_key, wk in history.items():
        print(f"  {wk_key}: {wk['label']} ({len(wk['cards'])} cards)")

    if args.dry:
        print("\n--dry flag set. Sample row:")
        print(json.dumps(rows[0], indent=2, ensure_ascii=False))
        return

    load_to_bigquery(rows)


if __name__ == "__main__":
    main()