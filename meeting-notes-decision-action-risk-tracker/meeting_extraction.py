"""
meeting_extraction.py

Shared logic for the meeting notes -> action items pipeline:
- The Claude API extraction schema and prompt (used by both the CLI script
  and the Streamlit app, so they always stay in sync)
- Excel export/append logic for the shared meeting_tracker.xlsx file

Importing this module has no side effects -- it just defines constants and
functions for extract_action_items.py and app.py to use.
"""

import os
import re
import difflib
from datetime import datetime
import anthropic
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule

MODEL = "claude-sonnet-4-6"  # solid default for extraction tasks; swap for
                            # claude-haiku-4-5-20251001 if you want cheaper/faster

# Dedup matching: meeting transcripts often re-raise the same decision/risk/
# action item week over week, but the LLM's phrasing of it won't be
# byte-identical (or even very close, character-for-character) across runs.
# A single character-level similarity check misses heavier paraphrasing
# ("we will not compress the QA window" vs "decided not to compress QA
# testing"), so we combine two signals and flag a duplicate if EITHER fires:
#   1. Character-level similarity (difflib) -- catches close rewordings
#   2. Significant-word overlap (Jaccard on non-stopword tokens) -- catches
#      heavier rewrites that still share the same key nouns/verbs
SEQUENCE_SIMILARITY_THRESHOLD = 0.72
WORD_OVERLAP_THRESHOLD = 0.38

_STOPWORDS = {
    "a", "an", "the", "is", "are", "was", "were", "be", "been", "being", "to",
    "of", "in", "on", "at", "for", "with", "by", "from", "and", "or", "but",
    "if", "not", "it", "this", "that", "will", "would", "could", "should",
    "may", "might", "we", "they", "has", "have", "had", "do", "does", "did",
    "than", "then", "so", "as", "no", "yes", "still", "further", "past",
    "due", "based", "into", "out", "up", "down", "over", "under",
}

EXCEL_PATH = "meeting_tracker.xlsx"

# Column layout for each tab. "Date Complete" and "Mitigation Plan" are left
# blank on extraction -- transcripts don't state them, and the project's
# anti-hallucination principle means we don't want Claude guessing at them.
# Fill those in manually (in Excel, or in the Streamlit review step) as work
# progresses.
DECISIONS_COLUMNS = ["Meeting Date", "Decision", "Context", "Comments"]
# Priority comes before Status/Date Complete since it's set at extraction
# time; Status and Date Complete are for manual tracking as work progresses.
ACTION_ITEMS_COLUMNS = ["Meeting Date", "Task", "Owner", "Due Date", "Priority", "Status", "Date Complete", "Comments"]
# Severity and Impact are manual fields (the extraction schema doesn't
# capture them -- transcripts don't reliably state a formal severity/impact
# rating). Owner sits after Mitigation Plan by request. Status and Date
# Closed are for manual tracking, Comments for free-form notes.
RISKS_COLUMNS = ["Meeting Date", "Risk", "Severity", "Impact", "Mitigation Plan", "Owner", "Status", "Date Closed", "Comments"]

# Allowed values for the Status dropdown (Excel data validation) on each tab.
# Kept short and mutually exclusive so the column is actually useful for
# filtering/sorting, rather than accumulating free-text variations.
ACTION_ITEM_STATUS_OPTIONS = ["Not Started", "In Progress", "Blocked", "Complete"]
RISK_STATUS_OPTIONS = ["Open", "Mitigated", "Closed", "Accepted"]

# Status colors, used for conditional formatting on the Status column of each
# tab. Colors are deliberately shared where the meaning lines up (e.g. both
# "Blocked" and "Open" are a red/urgent state) so the color language is
# consistent across tabs even though the exact status words differ.
STATUS_FILL_COLORS = {
    "Not Started": "D9D9D9",  # neutral gray
    "In Progress": "FFF2CC",  # light yellow
    "Blocked": "F4CCCC",      # light red
    "Complete": "D9EAD3",     # light green
    "Open": "F4CCCC",         # light red
    "Mitigated": "FFF2CC",    # light yellow
    "Closed": "D9EAD3",       # light green
    "Accepted": "CFE2F3",     # light blue -- accepted risk, distinct from resolved
}

# Audit trail: one row per script run, so the tracker carries its own history
# of what happened on each run -- not just the accumulated data. "Notes"
# holds a short human-readable summary of anything skipped, so a reviewer
# can see at a glance whether a run had unusual activity worth checking.
NOTES_COLUMNS = ["Run Timestamp", "Meeting Date Processed", "Decisions Added",
                  "Decisions Skipped", "Action Items Added", "Action Items Skipped",
                  "Risks Added", "Risks Skipped", "Notes"]

# Explicit, literal column widths (Excel character units), one list per tab,
# paired 1:1 with that tab's COLUMNS list, applied directly with no
# proportional scaling. Kept as separate per-tab lists (rather than one
# shared name-keyed dict) because the same column name can need a different
# width on different tabs -- e.g. "Comments" is 36 on Risks but also 36 here
# on Decisions/Action Items in this version; keeping them as separate lists
# still allows them to diverge independently if that changes again.
DECISIONS_WIDTHS = [10, 36, 36, 36]
ACTION_ITEMS_WIDTHS = [10, 36, 16, 16, 10, 10, 10, 36]
RISKS_WIDTHS = [10, 36, 10, 36, 36, 10, 10, 10, 36]
NOTES_WIDTHS = [10, 10, 10, 10, 10, 10, 10, 10, 60]

# Floor so a column never renders unreadably thin, even if a future column
# is added without an explicit width.
MIN_COLUMN_WIDTH = 10

# Font applied throughout the entire workbook -- Calibri 11pt is Excel's own
# default, but set explicitly (rather than left implicit) so the workbook is
# consistent regardless of what created/edited it, including after a user's
# manual formatting changes get overwritten on the next run.
WORKBOOK_FONT_NAME = "Calibri"
WORKBOOK_FONT_SIZE = 11

HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
HEADER_FONT = Font(name=WORKBOOK_FONT_NAME, size=WORKBOOK_FONT_SIZE, bold=True, color="FFFFFF")
DATA_FONT = Font(name=WORKBOOK_FONT_NAME, size=WORKBOOK_FONT_SIZE)

# The schema below is what forces Claude's output into a predictable shape.
# `strict: true` guarantees the returned JSON matches this exactly --
# no markdown fences, no "Here's the JSON:" preamble, no missing fields.
EXTRACTION_TOOL = {
    "name": "record_meeting_extraction",
    "description": "Record the structured extraction of a meeting transcript.",
    "input_schema": {
        "type": "object",
        "properties": {
            "meeting_date": {
                "type": "string",
                "description": (
                    "The date of the meeting, as stated in the transcript "
                    "(e.g. from a header like 'Date: September 3, 2026'). "
                    "Use 'Not specified' if no date is stated anywhere."
                )
            },
            "summary": {
                "type": "string",
                "description": "2-3 sentence plain-language recap of the meeting."
            },
            "decisions": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "decision": {"type": "string"},
                        "context": {
                            "type": "string",
                            "description": "Why this decision was made / what it affects."
                        }
                    },
                    "required": ["decision", "context"],
                    "additionalProperties": False
                }
            },
            "action_items": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "task": {"type": "string"},
                        "owner": {
                            "type": "string",
                            "description": "Person responsible. Use 'Unassigned' if unclear."
                        },
                        "due_date": {
                            "type": "string",
                            "description": "Due date if stated, else 'Not specified'."
                        },
                        "priority": {
                            "type": "string",
                            "enum": ["high", "medium", "low"]
                        }
                    },
                    "required": ["task", "owner", "due_date", "priority"],
                    "additionalProperties": False
                }
            },
            "open_risks": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "risk": {"type": "string"},
                        "mitigation_plan": {
                            "type": "string",
                            "description": (
                                "Only include a mitigation plan if one was "
                                "explicitly stated in the transcript. "
                                "Otherwise use 'Not specified' -- don't infer one."
                            )
                        }
                    },
                    "required": ["risk", "mitigation_plan"],
                    "additionalProperties": False
                },
                "description": "Unresolved risks, blockers, or dependencies raised in the meeting."
            }
        },
        "required": ["meeting_date", "summary", "decisions", "action_items", "open_risks"],
        "additionalProperties": False
    },
    "strict": True
}

EXTRACTION_PROMPT_TEMPLATE = (
    "Extract the meeting date, decisions, action items, open risks, and a "
    "short summary from this meeting transcript. Be precise -- only include "
    "what's actually stated, don't infer names, dates, or mitigation plans "
    "that weren't said.\n\n"
    "For decisions specifically: scan the transcript for every distinct "
    "moment the group actually reaches or announces a decision -- these are "
    "often signaled by words like 'decision:', 'let's decide', 'agreed', "
    "'we're going with', or similar. Log EACH one as its own separate entry "
    "using the specific, concrete details stated at that moment (the actual "
    "date chosen, the actual option picked).\n\n"
    "This includes cases where a new decision resolves a conditional one "
    "from earlier in the SAME transcript -- for example, if the transcript "
    "shows the group first saying 'if X isn't fixed by Friday, we'll push "
    "the date' and later in the same conversation actually saying 'decision: "
    "we're pushing the cutover to the 22nd', these are TWO separate "
    "decisions and both must appear as separate entries. Do not collapse "
    "the concrete new decision into a restatement of the earlier "
    "conditional one, and do not describe the new decision using only the "
    "old conditional's wording -- use the new decision's own specific "
    "details (the actual date/option that was decided).\n\n"
    "For open_risks specifically: treat each distinct underlying cause or "
    "blocker as one risk item. Do not split a single risk into multiple "
    "entries just because it was mentioned more than once or discussed from "
    "different angles -- only create separate risk items when they have "
    "genuinely different causes, owners, or mitigation paths.\n\n"
    "Before calling the tool, first write a short numbered list in plain "
    "text of every distinct decision you found in the transcript (just the "
    "count and a one-line description of each) -- then call the tool with "
    "the full extraction. This helps you enumerate decisions accurately "
    "instead of merging related ones together."
    "{decision_hints_block}\n\n"
    "TRANSCRIPT:\n{transcript_text}"
)

# Explicit decision-language markers, found via regex rather than relying on
# the model to notice them unaided. In testing, prompt wording alone
# (multiple attempts) wasn't reliable at getting Claude to log a concrete new
# decision as its own entry when it resolved an earlier conditional one --
# the model tended to merge them. Surfacing literal textual anchors from the
# transcript grounds the extraction in what's actually written, rather than
# depending entirely on the model's holistic judgment of "what counts as a
# new decision." These are HINTS the model still has to interpret and
# paraphrase -- not used directly as output -- and any hint that turns out
# to just be a re-statement of something already logged is still correctly
# filtered out by the dedup logic afterward, so surfacing extra hints here
# carries little downside.
#
# The punctuation after the marker word is intentionally kept optional (to
# still catch phrasings like "decision now -- we will not compress..." or
# "decision logged." with no colon/dash), but "decision"/"decided" is
# excluded when immediately followed by "from" -- this specifically filters
# out backward-looking references like "the monitoring tool decision from
# last week", which describe a PAST decision being discussed, not a new one
# being announced, without over-broadly excluding genuine announcements.
DECISION_MARKER_PATTERN = re.compile(
    r"\b((?:[Dd]ecision(?!\s+from\b)(?:\s+logged)?|[Dd]ecided(?!\s+from\b))"
    r"\s*[:\-\u2014]?\s*[^.\n]*\.)"
)
MAX_DECISION_HINTS = 8
MIN_HINT_CONTENT_WORDS = 3  # below this, the hint is just an acknowledgment
                            # ("decision logged.") with no content of its
                            # own worth injecting as a standalone decision


def _hint_content_word_count(raw_hint: str) -> int:
    """Word count remaining after stripping the marker prefix -- used to
    filter out content-free acknowledgments like 'decision logged.' that
    would otherwise degenerate into a nonsense injected entry (e.g. just
    '.') since there's nothing substantive left to report."""
    stripped = re.sub(
        r"^(?:[Dd]ecision(?:\s+logged)?|[Dd]ecided)\s*[:\-\u2014]?\s*",
        "",
        raw_hint.strip()
    ).strip().rstrip(".")
    return len(stripped.split())


def _find_decision_marker_hints(transcript_text: str) -> list:
    """Regex-detect explicit decision-language snippets in the transcript,
    to ground the model's extraction in literal text rather than relying
    solely on its own judgment of what counts as a new decision."""
    matches = DECISION_MARKER_PATTERN.findall(transcript_text)
    # Preserve order, drop exact duplicates, cap length for prompt hygiene.
    seen = set()
    hints = []
    for m in matches:
        cleaned = " ".join(m.split())
        if cleaned.lower() in seen:
            continue
        if _hint_content_word_count(cleaned) < MIN_HINT_CONTENT_WORDS:
            continue  # e.g. bare "decision logged." -- nothing to report
        seen.add(cleaned.lower())
        hints.append(cleaned[:200])
        if len(hints) >= MAX_DECISION_HINTS:
            break
    return hints


def _build_decision_hints_block(transcript_text: str) -> str:
    hints = _find_decision_marker_hints(transcript_text)
    if not hints:
        return ""
    bullet_list = "\n".join(f"- \"{h}\"" for h in hints)
    return (
        "\n\nAs a starting point, a keyword search found these explicit "
        "decision-language snippets in the transcript -- make sure each one "
        "is accounted for as its own entry in your output (paraphrase them "
        "naturally; some may turn out to describe the same decision and "
        "should only appear once; some may be minor and not worth a "
        "separate entry -- use your judgment, but don't silently drop a "
        "genuinely new, concrete decision):\n" + bullet_list
    )


def _numeric_tokens(text: str) -> set:
    """Extract standalone numbers (e.g. day-of-month references like '22' or
    '22nd') for a lightweight same-event check between short hint fragments,
    which are often too short/differently-worded for the main word-overlap
    threshold to catch as duplicates of each other."""
    return set(re.findall(r"\b\d{1,4}\b", text))


def _hints_likely_same_decision(a: str, b: str) -> bool:
    """Lenient same-event check for short hint fragments specifically.
    Full sentences use _is_duplicate's stricter thresholds; short fragments
    like 'decision logged -- cutover moves to the 22nd.' vs 'Decision: we're
    pushing the cutover to the 22nd.' describe the same event but often
    fall below those thresholds due to how little text there is to match
    against. Sharing a specific number (day-of-month, etc.) is a strong
    signal of referring to the same decision; a lower word-overlap bar
    catches the rest."""
    a_nums, b_nums = _numeric_tokens(a), _numeric_tokens(b)
    if a_nums and b_nums and (a_nums & b_nums):
        return True
    return _word_overlap(a, b) >= 0.25


def _clean_decision_hint_for_injection(hint: str) -> str:
    """Turn a raw regex-captured hint into readable decision text: strip the
    leading marker word ('Decision:', 'decision logged —', 'Decided:'), then
    capitalize and ensure it ends with a period."""
    cleaned = re.sub(
        r"^(?:[Dd]ecision(?:\s+logged)?|[Dd]ecided)\s*[:\-\u2014]?\s*",
        "",
        hint.strip()
    ).strip()
    if not cleaned:
        cleaned = hint.strip()
    cleaned = cleaned[0].upper() + cleaned[1:] if cleaned else cleaned
    if not cleaned.endswith("."):
        cleaned += "."
    return cleaned


def _dedupe_hints_against_each_other(hints: list) -> list:
    """Multiple regex hints can describe the SAME real decision (e.g. both
    'Decision: we're pushing to the 22nd' and 'decision logged — cutover
    moves to the 22nd' refer to one event). Keep only the first of each
    near-duplicate group so we don't inject the same decision twice."""
    kept = []
    for h in hints:
        if not any(_hints_likely_same_decision(h, k) for k in kept):
            kept.append(h)
    return kept


def _fill_missed_decisions(result: dict, transcript_text: str) -> dict:
    """Safety net: if an explicit decision-marker hint from the transcript
    isn't reflected anywhere in Claude's own decisions list, inject it
    directly as its own decision entry.

    In testing, prompt wording alone -- across several increasingly direct
    attempts -- wasn't reliably enough to stop Claude from merging a new
    concrete decision into a restatement of an earlier conditional one it
    resolved (e.g. 'we'll push the date if X isn't fixed' + 'decision: we're
    pushing to the 22nd' collapsing into just the first). Rather than keep
    tuning the prompt indefinitely, this backstops the one guarantee that
    actually matters -- an explicit decision statement is never silently
    dropped -- with deterministic code instead of model judgment.
    """
    raw_hints = _find_decision_marker_hints(transcript_text)
    unique_hints = _dedupe_hints_against_each_other(raw_hints)

    existing_decision_texts = [d["decision"] for d in result.get("decisions", [])]
    for hint in unique_hints:
        already_covered = _is_duplicate(hint, existing_decision_texts) or any(
            _hints_likely_same_decision(hint, existing) for existing in existing_decision_texts
        )
        if already_covered:
            continue  # Claude already captured this one in its own words.
        injected_text = _clean_decision_hint_for_injection(hint)
        result.setdefault("decisions", []).append({
            "decision": injected_text,
            "context": (
                "Automatically captured from an explicit decision marker in "
                "the transcript that wasn't reflected in the model's own "
                "summary -- worth a quick human check."
            )
        })
        existing_decision_texts.append(injected_text)

    return result


def extract_from_transcript(transcript_text: str, client: anthropic.Anthropic = None) -> dict:
    """Call the Claude API to extract structured data from a transcript.

    Pass an existing client (e.g. a Streamlit-cached one) to reuse it,
    or omit it to create a new one that reads ANTHROPIC_API_KEY from env.
    """
    if client is None:
        client = anthropic.Anthropic()

    prompt = EXTRACTION_PROMPT_TEMPLATE.format(
        transcript_text=transcript_text,
        decision_hints_block=_build_decision_hints_block(transcript_text)
    )

    response = client.messages.create(
        model=MODEL,
        max_tokens=3000,
        tools=[EXTRACTION_TOOL],
        # Deliberately NOT forcing tool_choice to "tool" here. Forcing an
        # immediate tool call gives Claude no room to reason step-by-step
        # first, which in testing caused it to merge/conflate related
        # decisions (e.g. an old conditional decision and a new concrete one
        # resolving it) into a single entry. Allowing a short plain-text
        # enumeration first (see the prompt) measurably improved accuracy.
        # We still require max_tokens headroom for that reasoning text, and
        # the tool's strict schema still guarantees valid structure once the
        # tool is actually called.
        tool_choice={"type": "auto"},
        messages=[
            {
                "role": "user",
                "content": prompt
            }
        ]
    )

    # With tool_choice forcing this specific tool, the response's only
    # content block will be a tool_use block containing our validated JSON.
    for block in response.content:
        if block.type == "tool_use":
            return _fill_missed_decisions(block.input, transcript_text)

    raise RuntimeError("No tool_use block found in response -- unexpected.")


def _scaled_widths(widths: list, columns: list) -> list:
    """Apply each column's literal width directly (only a MIN_COLUMN_WIDTH
    floor for safety), with no proportional scaling. Falls back to a
    reasonable default if a columns list and its widths list don't line up
    (e.g. a column was added without updating its tab's widths list to
    match)."""
    if len(widths) != len(columns):
        return [max(MIN_COLUMN_WIDTH, 20)] * len(columns)
    return [max(MIN_COLUMN_WIDTH, w) for w in widths]


def _style_header(ws, columns: list, widths: list) -> None:
    """Apply header styling and the tab's scaled column widths to a fresh sheet."""
    scaled = _scaled_widths(widths, columns)
    for col_idx, (header, width) in enumerate(zip(columns, scaled), start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(vertical="center", wrap_text=False)
        ws.column_dimensions[get_column_letter(col_idx)].width = width
    ws.freeze_panes = "A2"


def _apply_column_widths(ws, columns: list, widths: list) -> None:
    """Set column widths from the tab's scaled weights. Applied on every run
    (not just sheet creation) so a tracker file made before a width change --
    or manually resized -- gets consistent widths again automatically."""
    for col_idx, width in enumerate(_scaled_widths(widths, columns), start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = width


def _apply_header_font(ws, columns: list) -> None:
    """Reapply the header font/fill to row 1. Applied on every run (not just
    sheet creation) so a tracker file made before this font was standardized
    -- or manually reformatted -- gets a consistent look again automatically."""
    for col_idx in range(1, len(columns) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL


def _apply_status_validation(ws, columns: list, status_options: list) -> None:
    """Add an Excel dropdown (data validation) to the Status column, so
    values stay consistent and the column is actually useful for
    filtering/sorting rather than accumulating free-text variations."""
    if "Status" not in columns or not status_options:
        return
    col_letter = get_column_letter(columns.index("Status") + 1)
    dv = DataValidation(
        type="list",
        formula1='"' + ",".join(status_options) + '"',
        allow_blank=True
    )
    dv.error = "Please choose a value from the dropdown."
    dv.errorTitle = "Invalid Status"
    ws.add_data_validation(dv)
    dv.add(f"{col_letter}2:{col_letter}1000")


def _apply_status_conditional_formatting(ws, columns: list, status_options: list) -> None:
    """Color the Status cell based on its value (see STATUS_FILL_COLORS), so
    open/blocked items are visually distinct from completed/closed ones at a
    glance. Only the status values valid for this specific tab are used."""
    if "Status" not in columns:
        return
    col_letter = get_column_letter(columns.index("Status") + 1)
    cell_range = f"{col_letter}2:{col_letter}1000"
    for status_value in status_options:
        color = STATUS_FILL_COLORS.get(status_value)
        if not color:
            continue
        fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
        ws.conditional_formatting.add(
            cell_range,
            CellIsRule(operator="equal", formula=[f'"{status_value}"'], fill=fill)
        )


def _apply_autofilter(ws) -> None:
    """Enable Excel's AutoFilter on the header row, covering all current data
    rows. Safe to reapply every run (unlike conditional formatting rules,
    setting auto_filter.ref just overwrites the previous single range rather
    than accumulating), so the filter always covers newly appended rows too."""
    if ws.max_row < 1 or ws.max_column < 1:
        return
    last_col = get_column_letter(ws.max_column)
    ws.auto_filter.ref = f"A1:{last_col}{ws.max_row}"


def _get_or_create_sheet(wb: Workbook, name: str, columns: list, widths: list,
                          status_options: list = None):
    """Return the named sheet, creating it with a styled header row if new.
    Column widths and header font are reapplied every time, even for
    existing sheets. widths is that tab's explicit literal-width list,
    paired 1:1 with columns. Pass status_options for a tab with a Status
    column, to set up its dropdown and color-coding -- applied only once at
    creation, since data validation/conditional formatting rules would
    otherwise accumulate duplicates on every run."""
    if name in wb.sheetnames:
        ws = wb[name]
        _apply_column_widths(ws, columns, widths)
        _apply_header_font(ws, columns)
        return ws
    ws = wb.create_sheet(name)
    _style_header(ws, columns, widths)
    if status_options:
        _apply_status_validation(ws, columns, status_options)
        _apply_status_conditional_formatting(ws, columns, status_options)
    return ws


def _normalize(text: str) -> str:
    """Lowercase and collapse whitespace for fair text comparison."""
    return " ".join(text.lower().split())


def _significant_words(text: str) -> set:
    """Lowercase word tokens with stopwords and very short words removed,
    used for word-overlap comparison."""
    words = re.findall(r"[a-z0-9']+", text.lower())
    return {w for w in words if w not in _STOPWORDS and len(w) > 2}


def _word_overlap(a: str, b: str) -> float:
    """Jaccard similarity of significant words between two strings."""
    wa, wb = _significant_words(a), _significant_words(b)
    if not wa or not wb:
        return 0.0
    return len(wa & wb) / len(wa | wb)


def _is_duplicate(candidate: str, existing_texts: list) -> bool:
    """True if candidate is a near-match of anything already in existing_texts.

    Flags a duplicate if EITHER the character-level similarity or the
    significant-word overlap clears its threshold -- see the module-level
    comment above SEQUENCE_SIMILARITY_THRESHOLD for why both are needed.
    """
    candidate_norm = _normalize(candidate)
    for existing in existing_texts:
        existing_norm = _normalize(existing)
        seq_ratio = difflib.SequenceMatcher(None, candidate_norm, existing_norm).ratio()
        if seq_ratio >= SEQUENCE_SIMILARITY_THRESHOLD:
            return True
        if _word_overlap(candidate_norm, existing_norm) >= WORD_OVERLAP_THRESHOLD:
            return True
    return False


def _is_duplicate_action_item(task: str, owner: str, existing_pairs: list) -> bool:
    """True if (task, owner) matches an existing (task, owner) pair.

    Owner must match (case-insensitive, exact) -- reassigning the same task
    to a different person is a genuinely different action item, not a
    duplicate. Task text is compared with the same hybrid similarity used
    elsewhere, since re-logged tasks are rarely worded identically.
    """
    owner_norm = _normalize(owner)
    for existing_task, existing_owner in existing_pairs:
        if _normalize(existing_owner) != owner_norm:
            continue
        if _is_duplicate(task, [existing_task]):
            return True
    return False


def _existing_column_values(ws, col_index: int) -> list:
    """Read all existing values from one column of a sheet (skipping the header row)."""
    values = []
    for row in ws.iter_rows(min_row=2, min_col=col_index, max_col=col_index, values_only=True):
        if row[0]:
            values.append(str(row[0]))
    return values


def build_workbook(result: dict, path: str = EXCEL_PATH) -> tuple:
    """Load the existing tracker (or create a new one) and append result's
    rows into it, skipping rows that closely match something already logged
    in a prior meeting (see SEQUENCE_SIMILARITY_THRESHOLD / WORD_OVERLAP_THRESHOLD).

    Returns (workbook, report) where report is a dict like:
        {
            "decisions": {"added": 2, "skipped": ["Renew monitoring tool contract"]},
            "action_items": {"added": 1, "skipped": [...]},
            "risks": {"added": 1, "skipped": [...]},
        }
    Caller is responsible for saving the workbook (to disk, or to a BytesIO
    buffer for a download button).
    """
    meeting_date = result["meeting_date"]
    report = {
        "decisions": {"added": 0, "skipped": []},
        "action_items": {"added": 0, "skipped": []},
        "risks": {"added": 0, "skipped": []},
    }

    if os.path.exists(path):
        wb = load_workbook(path)
    else:
        wb = Workbook()
        # Remove the default blank sheet openpyxl creates -- we add our own below.
        default_sheet = wb.active
        wb.remove(default_sheet)

    decisions_ws = _get_or_create_sheet(wb, "Decisions", DECISIONS_COLUMNS, DECISIONS_WIDTHS)
    action_items_ws = _get_or_create_sheet(
        wb, "Action Items", ACTION_ITEMS_COLUMNS, ACTION_ITEMS_WIDTHS,
        status_options=ACTION_ITEM_STATUS_OPTIONS
    )
    risks_ws = _get_or_create_sheet(
        wb, "Risks", RISKS_COLUMNS, RISKS_WIDTHS, status_options=RISK_STATUS_OPTIONS
    )

    # --- Decisions: dedup on the "Decision" column (index 2) ---
    existing_decisions = _existing_column_values(decisions_ws, col_index=2)
    for d in result["decisions"]:
        if _is_duplicate(d["decision"], existing_decisions):
            report["decisions"]["skipped"].append(d["decision"])
            continue
        # Comments (last column) is left blank for manual notes.
        decisions_ws.append([meeting_date, d["decision"], d["context"], ""])
        existing_decisions.append(d["decision"])  # so duplicates within the same run are also caught
        report["decisions"]["added"] += 1

    # --- Action items: dedup on task text, but only when the owner also
    # matches -- same task reassigned to a different person is a genuinely
    # different item, not a duplicate. ---
    existing_tasks = _existing_column_values(action_items_ws, col_index=2)
    existing_owners = _existing_column_values(action_items_ws, col_index=3)
    existing_task_owner_pairs = list(zip(existing_tasks, existing_owners))
    for a in result["action_items"]:
        if _is_duplicate_action_item(a["task"], a["owner"], existing_task_owner_pairs):
            report["action_items"]["skipped"].append(f"{a['task']} — {a['owner']}")
            continue
        # Status defaults to "Not Started" for a freshly-logged item; Status,
        # "Date Complete", and "Comments" (last three columns) are otherwise
        # intentionally left for manual tracking as work progresses.
        action_items_ws.append(
            [meeting_date, a["task"], a["owner"], a["due_date"], a["priority"], "Not Started", "", ""]
        )
        existing_task_owner_pairs.append((a["task"], a["owner"]))
        report["action_items"]["added"] += 1

    # --- Risks: dedup on the "Risk" column (index 2) ---
    existing_risks = _existing_column_values(risks_ws, col_index=2)
    for r in result["open_risks"]:
        if _is_duplicate(r["risk"], existing_risks):
            report["risks"]["skipped"].append(r["risk"])
            continue
        # Severity, Impact, and Owner are left blank -- the extraction schema
        # doesn't capture them (transcripts don't reliably state a formal
        # severity/impact rating or a risk owner), consistent with the same
        # "leave it for human judgment rather than infer" principle used
        # elsewhere. Status defaults to "Open" for a freshly-logged risk.
        risks_ws.append([
            meeting_date, r["risk"], "", "", r["mitigation_plan"], "", "Open", "", ""
        ])
        existing_risks.append(r["risk"])
        report["risks"]["added"] += 1

    # --- Notes: one audit-trail row per run, regardless of whether anything
    # was added or skipped. This gives the workbook its own run history --
    # useful for spotting an unusual run (e.g. a lot of skips) after the
    # fact, without needing to cross-reference console output or screenshots. ---
    notes_ws = _get_or_create_sheet(wb, "Notes", NOTES_COLUMNS, NOTES_WIDTHS)
    notes_parts = []
    for tab_name, key in [("Decisions", "decisions"), ("Action Items", "action_items"), ("Risks", "risks")]:
        skipped = report[key]["skipped"]
        if skipped:
            notes_parts.append(f"{tab_name} skipped: " + "; ".join(skipped))
    notes_text = " | ".join(notes_parts) if notes_parts else "No items skipped."
    notes_ws.append([
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        meeting_date,
        report["decisions"]["added"], len(report["decisions"]["skipped"]),
        report["action_items"]["added"], len(report["action_items"]["skipped"]),
        report["risks"]["added"], len(report["risks"]["skipped"]),
        notes_text,
    ])

    # Keep long cells readable with top-aligned, wrapped text, and apply the
    # workbook's standard font to every data cell -- including on sheets
    # that existed before this font was standardized, so a run always
    # normalizes the whole file rather than just the rows it touches.
    for ws in (decisions_ws, action_items_ws, risks_ws, notes_ws):
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = Alignment(vertical="top", wrap_text=True)
                cell.font = DATA_FONT
        _apply_autofilter(ws)

    return wb, report


def append_to_excel(result: dict, path: str = EXCEL_PATH) -> dict:
    """Append result's rows into the tracker at `path`, saving to disk.
    Returns the same report dict described in build_workbook().
    """
    wb, report = build_workbook(result, path=path)
    wb.save(path)
    return report
