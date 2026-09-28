# Meeting Notes → Decision, Action Items & Risks Tracker

A tool that turns raw meeting transcripts into a structured, running Excel tracker — decisions, owned action items with due dates, and flagged risks with mitigation plans — using Claude's API. Built as both a CLI script and a Streamlit UI, both feeding the same tracker file, with duplicate detection so re-discussing the same open item in a later meeting doesn't create a second entry for it.

https://github.com/user-attachments/assets/f53650e2-1d5e-4475-b346-bf7084766099

## Why I built this

For Process Optimization or Technical Delivery initiatives, a recurring time sink is turning meeting notes into clear, trackable follow-ups: who owns what, by when, and what risks need to be escalated — and keeping that up to date across a series of weekly syncs on the same project. This project automates that translation step using Claude, while keeping a human in the loop to review and edit before anything is saved, and handling the reality that the same risk or decision often gets discussed again the following week.

## What it does

1. **Input**: paste or upload a meeting transcript (`.txt`) — via the CLI script or the Streamlit UI
2. **Extract**: Claude's API pulls out the meeting date, a summary, decisions, action items (task/owner/due date/priority), and open risks (with mitigation plan, if one was stated) — enforced by a strict JSON schema via tool use, so the output is always predictable and parseable
3. **Deduplicate**: before anything is written, each new item is checked against everything already logged in the tracker. Since Claude's phrasing of a recurring risk or decision won't be identical week to week, matching uses fuzzy text similarity (not exact string match) — so a re-raised item is recognized as already-logged and skipped, while a genuinely new item (even on a similar topic) still gets added. Action items are matched on task *and* owner together, so reassigning the same task to a different person is correctly treated as a new item, not a duplicate.
4. **Review** (Streamlit only): every extracted field is editable before it's saved — the AI drafts, a human confirms
5. **Save**: append the new (non-duplicate) rows into `meeting_tracker.xlsx`, a single workbook with four tabs that accumulates across every meeting:
   - **Decisions**: Meeting Date, Decision, Context, Comments
   - **Action Items**: Meeting Date, Task, Owner, Due Date, Priority, Status, Date Complete, Comments
   - **Risks**: Meeting Date, Risk, Severity, Impact, Mitigation Plan, Owner, Status, Date Closed, Comments
   - **Notes**: an audit trail — one row per script run, with a timestamp, the meeting date processed, added/skipped counts per category, and a plain-text summary of anything skipped, so the tracker carries its own run history, not just the accumulated data

   `Status` on Action Items and Risks is a real Excel dropdown, color-coded by value (red for Blocked/Open, yellow for In Progress/Mitigated, green for Complete/Closed, blue for Accepted). `Owner`, `Severity`, `Impact`, `Status`, `Date Complete`/`Date Closed`, and `Comments` are all left blank on extraction for a human to fill in as work progresses — none of them are inferred by the model. Every tab has AutoFilter enabled and uses fixed, explicit column widths (rather than auto-fit) so specific columns stay exactly as wide as intended; wider tabs like Risks may need horizontal scrolling on smaller screens as a result. The whole workbook uses a consistent Calibri 11pt font throughout, self-healing to correct formatting/widths on every run even if a tab was previously edited by hand. The CLI script and the Streamlit app both write to the same file. The Streamlit app also supports posting to Slack via webhook or downloading the raw JSON.

## Tech stack

- **Claude API** (Anthropic) — extraction via tool use with a strict JSON schema
- **openpyxl** — formatted, multi-tab Excel output with append-on-run behavior
- **Streamlit** — UI
- **Slack Incoming Webhooks** — optional output integration
- **Python**

## Key design decisions

- **Strict tool-use schema** instead of freeform prompting: guarantees the model's output always matches the expected structure (no markdown fences, no missing fields, no parsing edge cases). Every nested object explicitly sets `additionalProperties: false`, which the API now requires for strict schemas.
- **Shared extraction module** (`meeting_extraction.py`): the schema, prompt, and Excel/dedup logic live in one place, imported by both the CLI script and the Streamlit app. This guarantees the two interfaces can never drift out of sync with each other — a problem that already bit this project once before the refactor.
- **Fuzzy-match deduplication, not exact match**: an LLM won't reproduce identical wording for the same underlying risk or decision across separate meetings, so exact-string dedup would miss almost every real repeat. Matching combines character-level similarity with significant-word overlap, and flags a duplicate if either signal is strong enough — tuned against real paraphrased examples (e.g. "we will not compress the QA window" vs. "decided not to compress QA testing") rather than guessed at.
- **Granular risk extraction, not consolidated**: closely related risks (e.g. a technical root cause, a downstream QA blocker, and a launch-date impact) are kept as separate line items rather than merged into one vague entry, since each may need a different owner or follow-up — consolidating would lose actionable detail a TPM actually needs.
- **Human review step before sending** (Streamlit): AI extraction isn't perfect, especially with cross-talk or vague ownership in real meetings — this keeps a human as the final check, which matters more the more this gets used on real, high-stakes meetings.
- **Explicit anti-hallucination instructions**: the model is told not to infer owners, dates, or mitigation plans that weren't actually stated — "Date Complete" and "Mitigation Plan" are left blank on extraction by design rather than guessed at, and filled in manually as work progresses.
- **Append-only, never update-in-place**: once a row exists in the tracker, later runs will only skip it (if it's a re-raised duplicate) or leave it untouched — never edit it. This protects manual progress tracking (Status, Date Complete, Date Closed) from ever being silently overwritten by a future extraction run, at the cost of the tool never being able to auto-detect "this risk sounds resolved now" from a later transcript. That inference is left as a human judgment call, consistent with the project's human-in-the-loop philosophy.
- **Fixed column widths over auto-fit-to-screen**: an earlier version auto-scaled every tab's columns to fit one screen without horizontal scrolling. That works well until specific columns need specific widths (e.g. the same "Comments" column needing a different width depending on how many other columns share its tab) — at that point auto-scaling and precise control are in tension, and precise control won. Wider tabs like the 9-column Risks tab now may need horizontal scrolling on smaller screens as a result.
- **Hybrid regex + LLM extraction, not pure prompting**: explicit decision-language markers (like "Decision:") are detected in the transcript with regex and surfaced to the model as anchors it must account for, rather than relying entirely on the model to notice them unaided. This is a deliberate choice to backstop a specific guarantee (an explicit decision statement shouldn't be silently dropped) with deterministic code where prompting alone proved unreliable — see Known Limitations below for where this still has a gap.

## Running it locally

```bash
git clone https://github.com/kavita-s12/tech-program-management.git
cd tech-program-management/meeting-notes-decision-action-risk-tracker
pip install -r requirements.txt
export ANTHROPIC_API_KEY=your_key_here      # Mac/Linux
# $env:ANTHROPIC_API_KEY="your_key_here"    # Windows PowerShell

# CLI version -- appends into meeting_tracker.xlsx
python extract_action_items.py sample_meeting_transcript.txt

# Streamlit UI -- same tracker, plus review/edit, Slack, and JSON export
streamlit run app.py
```

## Project structure

```
.
├── app.py                          # Streamlit UI
├── extract_action_items.py         # CLI version of the extraction logic
├── meeting_extraction.py           # Shared schema, prompt, and Excel/dedup logic
├── demo.gif                        # Demo recording of the Streamlit workflow
├── sample_meeting_transcript.txt   # Test input: meeting 1
├── sample_meeting_transcript_2.txt # Test input: meeting 2 (re-raises several
│                                    #   meeting 1 items in different wording,
│                                    #   to demonstrate dedup)
├── requirements.txt
└── README.md
```

## Known limitations

- **A new decision that resolves an earlier conditional one can be merged into the old decision's wording rather than logged as its own entry.** For example, if a meeting first says "if X isn't fixed by Friday, we'll push the date" and later in the same or a follow-up meeting says "decision: we're pushing the date to the 22nd," the model sometimes captures only the original conditional decision, not the concrete new outcome. I tried several fixes for this — stronger prompt instructions, giving the model reasoning room before its structured output, and a regex-based hint system that surfaces literal "Decision:"-marked text as anchors — and while each improved the extraction's behavior, none fully eliminated this specific edge case in live testing. Rather than keep tuning indefinitely, I documented it here and left it in: it's a legitimate example of why the Streamlit review step exists — a human catches what the model's judgment call misses. I consider knowing when to stop iterating and rely on a human checkpoint instead to be a reasonable engineering call for this kind of tool, not just an unfixed bug.

## Possible extensions

- Jira/Asana ticket creation instead of (or alongside) Slack
- Ingest transcripts directly from Zoom/Teams export or API
- Trend view across meetings — e.g. how long a given risk has stayed open
- Auth + persistent storage for team-wide use
- Semantic (embedding-based) dedup instead of text-similarity heuristics, for even more robust matching across heavily reworded repeats

## Try it

Run `sample_meeting_transcript.txt` first, then `sample_meeting_transcript_2.txt` — the second is a realistic one-week-later follow-up meeting that revisits several items from the first (a resolved decision that gets referenced again, an evolving risk, a reassigned task) alongside genuinely new ones. Open the resulting `meeting_tracker.xlsx` and you should see: no duplicate rows for the items that were already logged, every genuinely new decision, action item, and risk from the second meeting present, a working Status dropdown with color-coding on the Action Items and Risks tabs, AutoFilter on every tab's header row, and a Notes tab with one audit row per run.
