#!/usr/bin/env python3
"""
CES Weekly Update -- PORTABLE DEMO generator
=============================================
Builds a completely standalone, shareable-by-email HTML file that showcases
every feature of the real Weekly Update page (week picker, cross-week search,
section grouping, Appreciate/Feature/Escalate modal, View Report button)
WITHOUT exposing any real CES data.

This script only ever READS template.html and history.json, purely to
borrow the exact same rendering engine and to copy the *shape* of the data
(number of weeks, section mix, which cards have a report link / next steps).
It never modifies, overwrites, or re-publishes template.html/history.json/
update.py. Output goes to index.html at the repo root -- root is a GitHub
Pages compatible source (Settings -> Pages -> Source: "main" branch, "/"
root). It'll be live at https://<you>.github.io/<repo>/ with zero extra
config.

CAUTION: Pages serving from the repo root publishes EVERY file in the repo,
not just index.html -- including history.json (real CES content) and
update.py. Only push this repo to Pages if it's fine for those files to be
publicly fetchable too, or gitignore/remove them from the branch you deploy.

Usage:
    python3 generate_demo.py
"""

import json
import sys
from pathlib import Path

HERE         = Path(__file__).parent
REAL_HISTORY = HERE / "history.json"
OUTPUT       = HERE / "index.html"

DEMO_REPORT_URL = "https://www.google.com"
DEMO_NAME       = "Jane Doe, Director of Analytics"
DEMO_EMAIL      = "demo@example.com"

sys.path.insert(0, str(HERE))
from update import build_html, validate_html  # noqa: E402  (reuse, never modify)

LOREM_TITLES = [
    "Lorem Ipsum Dolor Sit Amet Analysis",
    "Consectetur Adipiscing Elit Dashboard",
    "Sed Do Eiusmod Tempor Initiative",
    "Ut Labore Et Dolore Magna Report",
    "Ut Enim Ad Minim Veniam Workflow",
    "Quis Nostrud Exercitation Deep Dive",
    "Duis Aute Irure Dolor Scorecard",
    "Excepteur Sint Occaecat Reporting",
    "Cupidatat Non Proident Enhancement",
    "Sunt In Culpa Qui Officia Review",
]

LOREM_SUMMARY = (
    "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod "
    "tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim "
    "veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea "
    "commodo consequat."
)

LOREM_VALUE = (
    "Duis aute irure dolor in reprehenderit in voluptate velit esse cillum "
    "dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non "
    "proident, sunt in culpa qui officia deserunt mollit anim id est laborum."
)

LOREM_NS = (
    "Continue lorem ipsum dolor sit amet and align with stakeholders on next "
    "steps before broader rollout"
)


def _anonymize_card(card: dict, idx: int) -> dict:
    """Replace real content with placeholder Lorem Ipsum, keep the shape."""
    return {
        "section": card["section"],
        "title": LOREM_TITLES[idx % len(LOREM_TITLES)],
        "text": [LOREM_SUMMARY, LOREM_VALUE],
        "ns": LOREM_NS if card.get("ns") else "",
        "report": DEMO_REPORT_URL if card.get("report") else "",
    }


def build_demo_history() -> dict:
    """Mirror the real history.json's *shape* (weeks/sections/report/ns mix)
    with zero real content -- every string is generic Lorem Ipsum."""
    real_history = json.loads(REAL_HISTORY.read_text(encoding="utf-8"))
    demo_history = {}
    for week_key, week in real_history.items():
        demo_history[week_key] = {
            "label": week["label"],
            "cards": [
                _anonymize_card(card, i) for i, card in enumerate(week["cards"])
            ],
        }
    return demo_history


def anonymize_header(html: str) -> str:
    """Scrub the real name/email baked into the template so the demo has no
    personal contact info in it either."""
    html = html.replace("Victor Chowdhury, Director of Analytics", DEMO_NAME)
    html = html.replace("victor.chowdhury@walmart.com", DEMO_EMAIL)
    return html


def main():
    if not REAL_HISTORY.exists():
        sys.exit(f"ERROR: {REAL_HISTORY} not found (needed to copy the data shape)")

    (HERE / ".nojekyll").touch()  # tell GitHub Pages to skip Jekyll processing

    demo_history = build_demo_history()
    html = build_html(demo_history)          # same engine, real template.html, untouched on disk
    html = anonymize_header(html)
    validate_html(html)

    OUTPUT.write_text(html, encoding="utf-8")
    print(f"Demo written to {OUTPUT} ({len(html) // 1024}KB)")
    print("100% standalone -- no BigQuery, no live data, safe to email/share externally.")
    print("GitHub Pages ready: enable Settings -> Pages -> Source: main / (root)")


if __name__ == "__main__":
    main()
