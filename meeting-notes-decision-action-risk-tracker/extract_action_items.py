"""
extract_action_items.py

Reads a meeting transcript and uses the Claude API (tool use with a strict
JSON schema) to extract decisions, action items, open risks, and a summary.
Appends the results into a running Excel tracker (meeting_tracker.xlsx) with
three tabs: Decisions, Action Items, Risks.

Setup:
    pip install anthropic openpyxl
    $env:ANTHROPIC_API_KEY="your_key_here"      (PowerShell)
    export ANTHROPIC_API_KEY=your_key_here      (Mac/Linux)

Usage:
    python extract_action_items.py sample_meeting_transcript.txt

Each run appends new rows to meeting_tracker.xlsx in the current folder.
Existing rows (and any manual edits you've made, like marking an action item
complete) are preserved -- only new rows are added.

The extraction schema and Excel logic live in meeting_extraction.py, shared
with app.py (the Streamlit UI), so both stay in sync.
"""

import sys
from meeting_extraction import extract_from_transcript, append_to_excel, EXCEL_PATH


def pretty_print(result: dict) -> None:
    print(f"\nMeeting Date: {result['meeting_date']}")

    print("\n=== SUMMARY ===")
    print(result["summary"])

    print("\n=== DECISIONS ===")
    for d in result["decisions"]:
        print(f"- {d['decision']}")
        print(f"    context: {d['context']}")

    print("\n=== ACTION ITEMS ===")
    for a in result["action_items"]:
        print(f"- [{a['priority'].upper()}] {a['task']}")
        print(f"    owner: {a['owner']} | due: {a['due_date']}")

    print("\n=== OPEN RISKS ===")
    for r in result["open_risks"]:
        print(f"- {r['risk']}")
        print(f"    mitigation plan: {r['mitigation_plan']}")


def print_dedup_report(report: dict) -> None:
    print("\n=== TRACKER UPDATE ===")
    for tab_name, key in [("Decisions", "decisions"), ("Action Items", "action_items"), ("Risks", "risks")]:
        added = report[key]["added"]
        skipped = report[key]["skipped"]
        print(f"{tab_name}: {added} added" + (f", {len(skipped)} skipped as duplicates" if skipped else ""))
        for item in skipped:
            print(f"    (already logged) {item}")


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python extract_action_items.py <transcript_file>")
        sys.exit(1)

    with open(sys.argv[1], "r", encoding="utf-8") as f:
        transcript = f.read()

    result = extract_from_transcript(transcript)

    # Print human-readable version
    pretty_print(result)

    # Append into the running Excel tracker, skipping near-duplicate rows
    report = append_to_excel(result)
    print_dedup_report(report)
    print(f"\nSaved to {EXCEL_PATH} (Meeting Date: {result['meeting_date']})")
