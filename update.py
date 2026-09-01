#!/usr/bin/env python3
"""
CES Weekly Update Publisher
============================
Usage:
    python3 update.py --docx "/path/to/Weekly Update.docx" --week "Week of July 22, 2026" --key jul22

Arguments:
    --docx   Path to the Word doc  (default: OneDrive location)
    --week   Human label shown in the dropdown  (required)
    --key    Short JS-safe key, no spaces  (required, e.g. jul22)
    --dry    Print what would change without publishing

What this script does and does NOT touch:
    DOES:   Reads the Word doc, extracts cards, appends to history.json,
            injects history.json into template.html, publishes via share-puppy.
    NEVER:  Modifies template.html's CSS, JS, modal, or layout in any way.
            The template is treated as read-only source of truth.
"""

import argparse
import json
import re
import subprocess
import sys
from pathlib import Path

HERE       = Path(__file__).parent
TEMPLATE   = HERE / "template.html"
HISTORY_F  = HERE / "history.json"
DOCX_DEFAULT = Path.home() / "Library/CloudStorage/OneDrive-WalmartInc/Documents/Weekly Update/Weekly Update.docx"
OUTPUT     = Path("/tmp/ces-weekly-update-publish.html")

SECTION_HEADER_WORDS = {"OPS", "WFM", "CRM"}


# ── Docx parsing ─────────────────────────────────────────────────────────────

def parse_docx(docx_path: Path) -> list[dict]:
    """Extract cards from the Word doc. Returns list of card dicts."""
    try:
        from docx import Document
    except ImportError:
        sys.exit("ERROR: python-docx not installed. Run: uv pip install python-docx --index-url https://pypi.ci.artifacts.walmart.com/artifactory/api/pypi/external-pypi/simple --allow-insecure-host pypi.ci.artifacts.walmart.com")

    doc = Document(str(docx_path))

    # Collect hyperlinks: {rId: url}
    rels = {
        rel_id: rel._target
        for rel_id, rel in doc.part.rels.items()
        if "hyperlink" in rel.reltype.lower()
    }

    cards = []
    current_section = "OPS"
    block: list[tuple[str, object]] = []  # (text, paragraph) pairs for the card being built

    def flush():
        if block:
            card = _parse_card_block(list(block), current_section, rels)
            if card:
                cards.append(card)
        block.clear()

    for para in doc.paragraphs:
        text = para.text.strip()

        # Blank paragraph = card separator in the new (blank-line-delimited) doc layout
        if not text:
            flush()
            continue

        # Section header line
        if text in SECTION_HEADER_WORDS:
            flush()
            current_section = text
            continue

        block.append((text, para))

    flush()
    return cards


def _parse_card_block(lines: list[tuple[str, object]], section: str, rels: dict) -> dict | None:
    """Turn one blank-line-delimited block of paragraphs into a card dict.

    Supports two doc layouts:
      1. Legacy: single paragraph "Title: body text ... next steps ..."
      2. Current: "Title" paragraph, then separate "Summary:" / "Business Value:"
         (and optional "Next Steps:") paragraphs, each possibly wrapping onto
         following un-labeled paragraphs.
    """
    if not lines:
        return None

    title_line, title_para = lines[0]
    body_lines = lines[1:]

    # Legacy single-paragraph layout: "Title: body..." all in one paragraph.
    if not body_lines and ":" in title_line:
        colon = title_line.find(":")
        title = _clean_title(title_line[:colon].strip())
        body  = title_line[colon + 1:].strip()
        paras, next_steps = _split_body(body)
        return {
            "section": section,
            "title":   title,
            "text":    paras,
            "ns":      next_steps,
            "report":  _find_report_url(title_para, rels),
        }

    # No body at all and no colon -- can't make a card out of a lone title.
    if not body_lines:
        return None

    # Current layout: title paragraph + labeled Summary/Business Value paragraphs.
    title = _clean_title(title_line.rstrip(":").strip())
    label_patterns = {
        "summary": re.compile(r"^summary:?\s*(.*)$", re.IGNORECASE),
        "value":   re.compile(r"^business value:?\s*(.*)$", re.IGNORECASE),
        "ns":      re.compile(r"^next steps?:?\s*(.*)$", re.IGNORECASE),
    }
    buckets = {"summary": [], "value": [], "ns": []}
    current = "summary"  # default bucket until a label is seen
    report_url = ""

    for line, para in body_lines:
        for key, pat in label_patterns.items():
            m = pat.match(line)
            if m:
                current = key
                content = m.group(1).strip()
                if content:
                    buckets[key].append(content)
                break
        else:
            buckets[current].append(line)

        if not report_url:
            report_url = _find_report_url(para, rels)

    text_paras = [_join_lines(buckets["summary"])] if buckets["summary"] else []
    if buckets["value"]:
        text_paras.append(_join_lines(buckets["value"]))
    next_steps = _join_lines(buckets["ns"]).rstrip(".")

    return {
        "section": section,
        "title":   title,
        "text":    text_paras,
        "ns":      next_steps,
        "report":  report_url,
    }


def _clean_title(title: str) -> str:
    """Strip a trailing ", Status: In Progress"-style annotation off a title.

    Some weeks' doc authors tack a status label onto the same paragraph as
    the title (e.g. "FIXIT Operations Dashboard, Status: Work In Progress").
    That's noise, not part of the title -- every other card in every other
    week omits it, so strip it for consistency.
    """
    return re.sub(r",?\s*Status:\s*.+$", "", title, flags=re.IGNORECASE).strip()


def _join_lines(lines: list[str]) -> str:
    """Join wrapped/bulleted lines into one paragraph with sane punctuation.

    Some doc authors write Summary/Business Value as separate un-punctuated
    bullet lines (no doc list markers survive into plain paragraph text), so
    a naive " ".join() produces a run-on sentence. Insert a period between
    lines that don't already end in terminal punctuation.
    """
    parts = []
    for line in lines:
        line = line.strip()
        if not line:
            continue
        if parts and not parts[-1].endswith((".", "!", "?", ":")):
            parts[-1] += "."
        parts.append(line)
    return " ".join(parts)


def _split_body(body: str) -> tuple[list[str], str]:
    """Split body text into body paragraphs and a next-steps string."""
    ns_patterns = [
        r"Next steps? (?:are|is) to (.+)",
        r"Next steps?:\s*(.+)",
        r"Next step is to (.+)",
        r"Next step:\s*(.+)",
    ]
    next_steps = ""
    remaining = body

    for pattern in ns_patterns:
        m = re.search(pattern, remaining, re.IGNORECASE | re.DOTALL)
        if m:
            next_steps = m.group(1).strip().rstrip(".")
            remaining  = remaining[: m.start()].strip()
            break

    # Split remaining into multiple paragraphs at double newlines
    paras = [p.strip() for p in re.split(r"\n{2,}", remaining) if p.strip()]
    if not paras and remaining:
        paras = [remaining]

    return paras, next_steps


def _find_report_url(para, rels: dict) -> str:
    """Return a dashboard-report hyperlink found in the paragraph, or empty string.

    Only puppy.walmart.com links count as "View Report" targets. Other inline
    hyperlinks (e.g. Jira ticket references cited within the narrative) are
    intentionally ignored -- they aren't reports and shouldn't drive the
    "View Report" button.
    """
    try:
        xml = para._p.xml
        for rel_id, url in rels.items():
            if rel_id in xml and "puppy.walmart.com" in url:
                return url
    except Exception:
        pass
    return ""


# ── Build and validate ────────────────────────────────────────────────────────

def build_html(history: dict) -> str:
    """Inject history into the template. Template is NEVER modified."""
    template = TEMPLATE.read_text(encoding="utf-8")

    # Sanity-check the template has not been corrupted
    if "__HISTORY__" not in template:
        sys.exit(
            "ERROR: template.html is missing the __HISTORY__ placeholder.\n"
            "The template may have been accidentally overwritten.\n"
            "Restore it from: /Users/v0c003n/.code_puppy/puppy_share/v0c003n__ces-weekly-update__v9.html\n"
            "and re-run the setup step."
        )

    # The newest week must be first (it becomes CURRENT).
    # Sort by the label date — NOT by key string (string sort puts jul8 > jul22).
    from datetime import datetime
    def _label_date(k):
        try:
            return datetime.strptime(history[k]['label'], 'Week of %B %d, %Y')
        except Exception:
            return datetime.min
    weeks_ordered = sorted(history.keys(), key=_label_date, reverse=True)
    history_ordered = {k: history[k] for k in weeks_ordered}

    # Update CURRENT to the newest key
    newest_key = weeks_ordered[0]
    history_js = json.dumps(history_ordered, ensure_ascii=False, separators=(",", ":"))

    html = template.replace("__HISTORY__", history_js, 1)
    html = re.sub(r"var CURRENT='[^']*'", f"var CURRENT='{newest_key}'", html)

    # Update the <select> options to match new week list
    newest_label = history[newest_key]["label"]
    options = "\n".join(
        '<option value="{k}"{sel}>{label}{sfx}</option>'.format(
            k=k,
            sel=" selected" if k == newest_key else "",
            label=history[k]["label"],
            sfx=" (current)" if k == newest_key else "",
        )
        for k in weeks_ordered
    )
    html = re.sub(
        r'<select id="week-picker"[^>]*>.*?</select>',
        f'<select id="week-picker" class="week-picker" onchange="document.getElementById(\'search-input\').value=\'\';document.getElementById(\'search-clear\').style.display=\'none\';renderWeek(this.value)">\n{options}\n    </select>',
        html,
        flags=re.S,
    )

    return html


def validate_html(html: str) -> None:
    """Assert the output has all the structural markers we care about."""
    checks = {
        "__HISTORY__ not in output":         "__HISTORY__" not in html,
        "var HISTORY= present":              "var HISTORY=" in html,
        "renderWeek present":                "function renderWeek" in html,
        "renderSearch present":              "function renderSearch" in html,
        "search-input present":              'id="search-input"' in html,
        "openModal present":                 "function openModal" in html,
        "modal-bg present":                  "modal-bg" in html,
        "no </script> inside script block":  _script_block_clean(html),
        "week-picker select present":        'id="week-picker"' in html,
    }
    failures = [label for label, ok in checks.items() if not ok]
    if failures:
        sys.exit("VALIDATION FAILED:\n" + "\n".join(f"  - {f}" for f in failures))
    print(f"  All {len(checks)} validation checks passed.")


def _script_block_clean(html: str) -> bool:
    m = re.search(r"<script>(.*)</script>", html, re.S)
    if not m:
        return False
    return "</script>" not in m.group(1)


# ── Main ─────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Publish CES Weekly Update")
    parser.add_argument("--docx", default=str(DOCX_DEFAULT), help="Path to Weekly Update.docx")
    parser.add_argument("--week", required=True, help='e.g. "Week of July 22, 2026"')
    parser.add_argument("--key",  required=True, help="Short key, no spaces, e.g. jul22")
    parser.add_argument("--dry",  action="store_true", help="Preview without publishing")
    args = parser.parse_args()

    docx_path = Path(args.docx)
    if not docx_path.exists():
        sys.exit(f"ERROR: Word doc not found at {docx_path}")
    if not TEMPLATE.exists():
        sys.exit(f"ERROR: template.html not found at {TEMPLATE}")

    # Auto-adjust week label to previous Saturday
    from datetime import datetime, timedelta
    def _prev_saturday(label: str) -> str:
        try:
            dt = datetime.strptime(label, "Week of %B %d, %Y")
        except ValueError:
            return label  # already in correct form or custom label
        days_back = (dt.weekday() - 5) % 7   # Saturday = weekday 5
        sat = dt - timedelta(days=days_back)
        return sat.strftime("Week of %B %-d, %Y")

    adjusted = _prev_saturday(args.week)
    if adjusted != args.week:
        print(f"Week label adjusted to previous Saturday: {args.week!r} → {adjusted!r}")
        args.week = adjusted

    print(f"\nParsing {docx_path.name} ...")
    cards = parse_docx(docx_path)
    if not cards:
        sys.exit("ERROR: No cards extracted from the Word doc. Check the format.")

    print(f"Extracted {len(cards)} cards:")
    for c in cards:
        tag = f"[{c['section']}]"
        rpt = " [has report link]" if c["report"] else ""
        print(f"  {tag} {c['title'][:70]}{rpt}")

    # Load and update history
    history = json.loads(HISTORY_F.read_text(encoding="utf-8")) if HISTORY_F.exists() else {}
    if args.key in history:
        print(f"\nWARNING: key '{args.key}' already exists in history. It will be overwritten.")
        yn = input("Continue? [y/N] ").strip().lower()
        if yn != "y":
            sys.exit("Aborted.")

    history[args.key] = {"label": args.week, "cards": cards}

    print(f"\nBuilding HTML ...")
    html = build_html(history)

    print(f"Validating ...")
    validate_html(html)

    OUTPUT.write_text(html, encoding="utf-8")
    print(f"Output written to {OUTPUT} ({len(html) // 1024}KB)")

    if args.dry:
        print("\n--dry flag set. Skipping publish and history save.")
        print(f"Preview the output at: {OUTPUT}")
        return

    # Save history only after successful build+validate
    HISTORY_F.write_text(json.dumps(history, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"history.json updated ({len(history)} weeks)")

    print("\nPublishing to Puppy Pages ...")
    print(f"  Run: code-puppy 'Publish {OUTPUT} to ces-weekly-update slug for v0c003n'")
    print(f"\nDone! File ready at: {OUTPUT}")
    print("Tell Code Puppy: publish /tmp/ces-weekly-update-publish.html to ces-weekly-update")

    # Sync to BigQuery (best-effort — never blocks the publish flow)
    print("\nSyncing to BigQuery ...")
    sync_script = HERE / "sync_bigquery.py"
    result = subprocess.run([sys.executable, str(sync_script)], capture_output=True, text=True)
    print(result.stdout)
    if result.returncode != 0:
        print(f"WARNING: BigQuery sync failed (page was still published):\n{result.stderr}")
    else:
        print("BigQuery sync complete: wmt-d2-prod.ces_weekly_updates_v0c003n.updates")


if __name__ == "__main__":
    main()
