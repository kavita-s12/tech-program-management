"""
plan_synthesizer.py

Reads several team project plans (.xlsx), normalizes them into one task list,
then extracts:
  1. Dependencies          (within-team, cross-team, and unresolved)
  2. Timeline conflicts    (dependency timing, assumptions vs. upstream plans,
                            milestone date mismatches, gates after cutover,
                            holidays, change freezes)
  3. Resource conflicts    (people committed to too much at the same time)
  4. Risks                 (reported status, notes, and derived from the above)
and writes them to an Excel workbook with four tabs.

Deterministic rules do the date math. An OPTIONAL Claude pass (--use-claude)
resolves free-text dependencies the rules cannot match to a task.

Setup:
    pip install openpyxl anthropic

Usage:
    python plan_synthesizer.py sample_plans
    python plan_synthesizer.py sample_plans -o plan_review.xlsx
    python plan_synthesizer.py sample_plans --use-claude
"""

from __future__ import annotations

import argparse
import glob
import os
import re
import sys
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from functools import lru_cache
from datetime import date, datetime, timedelta

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

MODEL = "claude-sonnet-4-6"

# ----------------------------------------------------------------------------
# CONFIG - edit these for your own plans
# ----------------------------------------------------------------------------

# OPTIONAL. Any plan file that matches none of these is handled automatically: its team
# name comes from the file name and its dependency wording is matched without aliases.
# Add an entry only to give a team a friendlier name, a fixed ID prefix, or aliases.
# A plan file is assigned to the first team whose "match" text appears in the file name. "aliases" are words that, when they appear in a
# dependency description, point at that team's plan (e.g. "Atlas" = Core).
TEAM_CONFIG = [
    {"match": ["core"], "label": "Core (Atlas)", "prefix": "C", "is_core": True,
     "aliases": ["atlas", "core", "client"]},
    {"match": ["vendor", "cloudbridge"], "label": "Vendor (CloudBridge)", "prefix": "V",
     "is_core": False, "aliases": ["vendor", "cloudbridge"]},
    {"match": ["partner_a", "partner a", "payments"], "label": "Partner A (PayNet)",
     "prefix": "PA", "is_core": False, "aliases": ["paynet", "partner a"]},
    {"match": ["partner_b", "partner b", "security"], "label": "Partner B (Data & Security)",
     "prefix": "PB", "is_core": False, "aliases": ["partner b"]},
]

RUNTIME_TEAMS: list = list(TEAM_CONFIG)    # teams in use for this run (set by main)
NOTE_SPLIT = re.compile(r"\.(?=\s|$)\s*|;\s*")      # splits a notes cell into sentences
ID_LIKE = re.compile(r"[A-Za-z]{1,6}[-.]?\d+(?:\.\d+)*")

# Header names (lower-case) each field may appear under in a plan.
COLUMN_ALIASES = {
    "id": ["id", "task id", "task #", "milestone id", "#"],
    "name": ["task", "task name", "deliverable", "activity", "name"],
    "owner": ["owner", "responsible", "lead", "assignee"],
    "start": ["start", "start date", "planned start"],
    "end": ["end", "end date", "finish", "planned finish"],
    "duration": ["duration", "duration (working days)", "duration (days)"],
    "deps": ["depends on", "dependencies", "prerequisites", "blocked by", "predecessor"],
    "status": ["status", "rag"],
    "notes": ["notes", "comments", "risk notes"],
}

# Non-working days that matter to the plans. Add your own company holidays.
HOLIDAYS = {
    date(2026, 11, 26): "Thanksgiving",
    date(2026, 11, 27): "Day after Thanksgiving",
    date(2026, 12, 24): "Christmas Eve",
    date(2026, 12, 25): "Christmas Day",
    date(2027, 1, 1): "New Year's Day",
}

RESOURCE_THRESHOLD = 3          # this many concurrent commitments = conflict
MILESTONE_WORDS = ["cutover", "go-live", "go live", "launch"]
GATE_PATTERN = re.compile(r"sign-?\s?off|certification|compliance|approval|readiness", re.I)
RISK_WORDS = re.compile(
    r"risk|blocked|block\b|waiting|outstanding|delay|lead time|reduced|timeout|"
    r"freeze|critical|slip|issue", re.I)
BAD_STATUS = re.compile(r"amber|red|at risk|blocked|delayed|off track", re.I)
GOOD_STATUS = re.compile(r"^(green|on track)$", re.I)
EMPTY_DEPS = {"", "-", "—", "none", "n/a", "na"}

SEV_ORDER = {"High": 0, "Medium": 1, "Low": 2}

# ----------------------------------------------------------------------------
# Data classes
# ----------------------------------------------------------------------------


@dataclass
class Task:
    team: str
    tid: str
    name: str
    owner: str
    start: date | None
    end: date | None
    deps_raw: str
    status: str
    notes: str
    source_file: str
    is_core: bool = False
    orig_id: str = ""          # ID as written in the plan
    prefix: str = ""           # team prefix, used if the ID must be made unique

    @property
    def is_milestone(self) -> bool:
        return self.start is not None and self.start == self.end

    @property
    def window(self) -> str:
        if not self.start or not self.end:
            return "dates missing"
        return f"{self.start:%b %d} - {self.end:%b %d}"


@dataclass
class Dep:
    down: Task
    up: Task | None
    raw_text: str
    source: str                 # "Plan column", "Notes", "Claude"
    confidence: str             # High / Medium / Low / Unresolved
    hint_team: str | None = None
    assumed: date | None = None
    note: str = ""
    dep_id: str = ""

    @property
    def kind(self) -> str:
        if self.up is None:
            return "Unresolved"
        return "Cross-team" if self.up.team != self.down.team else "Within team"


# ----------------------------------------------------------------------------
# Text and date helpers
# ----------------------------------------------------------------------------

STOP = set(
    "a an the of for to on in by and or with from is are be was were until per as at "
    "into this that it its their our new final team plan task needs need needed assumes "
    "assume waiting blocked block complete completed available confirm".split())
SYNONYMS = {"pen": "penetration", "uat": "acceptance"}

MON = (r"(Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|June?|July?|Aug(?:ust)?|"
       r"Sept?(?:ember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)")
MONTHS = {m: i for i, m in enumerate(
    ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"], 1)}
DATE_A = re.compile(r"\b(\d{1,2})[-\s]" + MON + r"\b", re.I)                 # 23-Oct
DATE_B = re.compile(r"\b" + MON + r"\.?\s+(\d{1,2})(?:\s*[-–]\s*(\d{1,2}))?\b", re.I)  # Nov 16-27


def stem(w: str) -> str:
    for suf in ("ing", "ed", "s"):
        if w.endswith(suf) and len(w) - len(suf) >= 3:
            w = w[: -len(suf)]
            break
    if len(w) >= 4 and w[-1] == w[-2] and w[-1] not in "ls":
        w = w[:-1]
    return w


@lru_cache(maxsize=None)
def tokens(text: str) -> frozenset:
    out = set()
    for w in re.sub(r"[^a-z0-9]+", " ", text.lower()).split():
        w = SYNONYMS.get(w, w)
        if w.isdigit() or w in STOP or len(w) < 2:
            continue
        out.add(stem(w))
    return frozenset(out)


def month_num(s: str) -> int:
    return MONTHS[s[:3].lower()]


def resolve_year(month: int, day: int, ref: date) -> date:
    """Pick the year that puts this date at or after (ref - 60 days)."""
    d = date(ref.year, month, day)
    if d < ref - timedelta(days=60):
        d = date(ref.year + 1, month, day)
    return d


def extract_dates(text: str, ref: date):
    """Return [(start, end, (span_start, span_end))] for dates found in free text."""
    found = []
    for m in DATE_A.finditer(text):
        try:
            d = resolve_year(month_num(m.group(2)), int(m.group(1)), ref)
        except ValueError:
            continue
        found.append((m.start(), m.end(), d, d))
    for m in DATE_B.finditer(text):
        mo = month_num(m.group(1))
        try:
            d1 = resolve_year(mo, int(m.group(2)), ref)
            d2 = resolve_year(mo, int(m.group(3)), ref) if m.group(3) else d1
        except ValueError:
            continue
        found.append((m.start(), m.end(), d1, d2))
    found.sort()
    return [(a, b, (s, e)) for s, e, a, b in found]


def strip_spans(text: str, spans) -> str:
    for s, e in sorted(spans, reverse=True):
        text = text[:s] + " " + text[e:]
    return text


def weekdays_between(start: date, end: date):
    d = start
    while d <= end:
        if d.weekday() < 5:
            yield d
        d += timedelta(days=1)


def add_workdays(start: date, n: int) -> date:
    d = start
    while n > 0:
        d += timedelta(days=1)
        if d.weekday() < 5:
            n -= 1
    return d


def parse_cell_date(value, default_year: int):
    """Excel date, or text like 2026-11-02, 02-Nov-2026, 11/02/2026, 'Nov 2', '2-Nov'."""
    if value is None or value == "":
        return None, False
    if isinstance(value, datetime):
        return value.date(), False
    if isinstance(value, date):
        return value, False
    s = str(value).strip()
    for fmt in ("%Y-%m-%d", "%d-%b-%Y", "%m/%d/%Y", "%d-%B-%Y", "%b %d, %Y", "%B %d, %Y"):
        try:
            return datetime.strptime(s, fmt).date(), False
        except ValueError:
            pass
    for fmt in ("%b %d", "%B %d", "%d-%b", "%d %b"):
        try:
            d = datetime.strptime(s, fmt)
            return date(default_year, d.month, d.day), True
        except ValueError:
            pass
    return None, False


# ----------------------------------------------------------------------------
# Loading plans
# ----------------------------------------------------------------------------


def open_workbook(path: str):
    try:
        return load_workbook(path, data_only=True)
    except PermissionError:
        raise SystemExit(
            f"ERROR: cannot read {path} (permission denied).\n"
            f"Most likely the file is open in Excel, or OneDrive is still syncing it. "
            f"Close it in Excel (and check the Excel icon in the taskbar), wait for OneDrive "
            f"to finish syncing, then run again.")


def infer_year(files: list[str], fallback: int) -> int:
    years = Counter()
    for f in files:
        wb = open_workbook(f)
        for row in wb.worksheets[0].iter_rows(values_only=True):
            for v in row:
                if isinstance(v, (datetime, date)):
                    years[v.year] += 1
                elif isinstance(v, str):
                    for y in re.findall(r"\b(20\d\d)\b", v):
                        years[int(y)] += 1
    return years.most_common(1)[0][0] if years else fallback


GENERIC_NAME_WORDS = {"plan", "project", "team", "schedule", "roadmap", "wbs", "timeline",
                      "workplan", "work", "final", "draft", "copy", "and", "the"}


def auto_label(path: str) -> str:
    """Team name from a file name: '5_infra_extra_plan.xlsx' -> 'Infra Extra'."""
    stem_name = os.path.splitext(os.path.basename(path))[0]
    stem_name = re.sub(r"^\d+[\s._-]*", "", stem_name)
    words = [w for w in re.split(r"[\s._-]+", stem_name) if w]
    kept = [w for w in words
            if w.lower() not in GENERIC_NAME_WORDS and not re.fullmatch(r"v?\d+(\.\d+)*", w.lower())]
    words = kept or words
    label = " ".join(w if (w.isupper() and len(w) <= 4) else w.capitalize() for w in words)
    return label or os.path.splitext(os.path.basename(path))[0]


def configure_teams(cfgs: list[dict]) -> None:
    global RUNTIME_TEAMS
    RUNTIME_TEAMS = cfgs


def resolve_teams(files: list[str], core_keyword: str | None = None) -> list[dict]:
    """One team config per plan file: from TEAM_CONFIG if the file name matches, else automatic."""
    cfgs = []
    for f in files:
        name = os.path.basename(f).lower()
        base = next((c for c in TEAM_CONFIG if any(k in name for k in c["match"])), None)
        if base:
            cfg = dict(base)
            cfg["auto"] = False
        else:
            cfg = {"label": auto_label(f), "prefix": "", "is_core": False,
                   "aliases": [], "match": [], "auto": True}
        if core_keyword:
            cfg["is_core"] = core_keyword.lower() in name
        cfgs.append(cfg)
    used = {c["prefix"] for c in cfgs if c["prefix"]}
    for c in cfgs:
        if c["prefix"]:
            continue
        words = [w for w in re.split(r"\W+", c["label"]) if w]
        base = "".join(w[0] for w in words).upper()[:3]
        if len(base) < 2:
            base = re.sub(r"[^A-Za-z0-9]", "", c["label"])[:3].upper() or "T"
        p, n = base, 2
        while p in used:
            p, n = f"{base}{n}", n + 1
        used.add(p)
        c["prefix"] = p
    return cfgs


def finalize_aliases(cfgs: list[dict], tasks: list[Task]) -> None:
    """For automatically-detected teams, let distinctive words in the team name act as
    aliases in dependency text (e.g. 'Infra team' -> the Infra plan). A word is only used
    if no task name contains it and no other team name shares it."""
    name_tokens = set()
    for t in tasks:
        name_tokens |= tokens(t.name)
    owners = Counter()
    for c in cfgs:
        if c["auto"]:
            owners.update(set(re.findall(r"[a-z]+", c["label"].lower())))
    for c in cfgs:
        if not c["auto"]:
            continue
        al = [c["label"].lower()]
        for w in set(re.findall(r"[a-z]+", c["label"].lower())):
            if (len(w) >= 4 and w not in GENERIC_NAME_WORDS and owners[w] == 1
                    and not (tokens(w) & name_tokens)):
                al.append(w)
        c["aliases"] = al
    configure_teams(cfgs)


def uniquify_ids(tasks: list[Task], warnings: list[str]) -> None:
    """If different plans reuse the same task ID, show those tasks as PREFIX:ID."""
    owners = defaultdict(set)
    for t in tasks:
        owners[t.orig_id.lower()].add(t.team)
    clash = {k for k, v in owners.items() if len(v) > 1}
    renamed = 0
    for t in tasks:
        if t.orig_id.lower() in clash:
            t.tid = f"{t.prefix}:{t.orig_id}"
            renamed += 1
    if renamed:
        warnings.append(f"{renamed} tasks share an ID with a task in another plan; they are "
                        f"shown as PREFIX:ID (for example {next(t.tid for t in tasks if ':' in t.tid)}).")
    seen = Counter((t.team, t.orig_id.lower()) for t in tasks)
    for (team, i), n in seen.items():
        if n > 1:
            warnings.append(f"{team}: task ID '{i}' appears {n} times in the same plan; "
                            f"dependencies on it may point at the wrong row.")


def load_plan(path: str, cfg: dict, default_year: int, warnings: list[str]) -> list[Task]:
    ws = open_workbook(path).worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    header = [str(h).strip().lower() if h is not None else "" for h in rows[0]]

    colmap = {}
    for fld, aliases in COLUMN_ALIASES.items():
        for i, h in enumerate(header):
            if h in aliases:
                colmap[fld] = i
                break
    print(f"\n{os.path.basename(path)}  ->  team '{cfg['label']}'"
          + ("  (detected from file name)" if cfg.get("auto") else ""))
    print("  column mapping: " + ", ".join(
        f"{f}='{rows[0][i]}'" for f, i in colmap.items()))

    missing = [f for f in ("name", "start") if f not in colmap]
    if "end" not in colmap and "duration" not in colmap:
        missing.append("end or duration")
    if missing:
        raise SystemExit(f"ERROR: {path}: could not find column(s) for {missing}. "
                         f"Headers seen: {header}. Add aliases in COLUMN_ALIASES.")

    def cell(row, fld):
        i = colmap.get(fld)
        return row[i] if i is not None and i < len(row) else None

    tasks, auto_n = [], 0
    for r_idx, row in enumerate(rows[1:], start=2):
        name = cell(row, "name")
        if not name:
            continue
        auto_n += 1
        raw_id = cell(row, "id")
        if isinstance(raw_id, float) and raw_id.is_integer():
            raw_id = int(raw_id)
        tid = str(raw_id).strip() if raw_id not in (None, "") else f"{cfg['prefix']}-{auto_n}"
        start, yearless_s = parse_cell_date(cell(row, "start"), default_year)
        end, yearless_e = parse_cell_date(cell(row, "end"), default_year)
        dur = cell(row, "duration")
        if end is None and start and dur not in (None, ""):
            end = add_workdays(start, int(float(dur)) - 1)
        if start and end and end < start and yearless_e:
            end = end.replace(year=end.year + 1)
        if not start or not end:
            warnings.append(f"{cfg['label']} row {r_idx} ({tid}): missing/unreadable dates; "
                            f"excluded from date checks.")
        tasks.append(Task(
            team=cfg["label"], tid=tid, name=str(name).strip(),
            owner=str(cell(row, "owner") or "").strip(),
            start=start, end=end,
            deps_raw=str(cell(row, "deps") or "").strip(),
            status=str(cell(row, "status") or "").strip(),
            notes=str(cell(row, "notes") or "").strip(),
            source_file=os.path.basename(path), is_core=cfg.get("is_core", False),
            orig_id=tid, prefix=cfg["prefix"]))
    print(f"  {len(tasks)} tasks loaded"
          + ("  (no ID column: generated IDs)" if "id" not in colmap else ""))
    return tasks


# ----------------------------------------------------------------------------
# People
# ----------------------------------------------------------------------------


def build_people(tasks: list[Task]) -> dict[str, str]:
    """Map name variants (full name, unique first name) -> canonical full name."""
    owners = sorted({t.owner for t in tasks if t.owner and t.owner.lower() != "unassigned"})
    firsts = Counter(o.split()[0].lower() for o in owners)
    variants = {}
    for o in owners:
        variants[o.lower()] = o
        if firsts[o.split()[0].lower()] == 1:
            variants[o.split()[0].lower()] = o
    return variants


def persons_in(text: str, people: dict[str, str]) -> set[str]:
    found = set()
    for variant, full in people.items():
        if re.search(r"\b" + re.escape(variant) + r"\b", text, re.I):
            found.add(full)
    return found


# ----------------------------------------------------------------------------
# Dependency extraction
# ----------------------------------------------------------------------------


def find_hint(text: str):
    hit, cleaned = None, text
    for cfg in RUNTIME_TEAMS:
        for a in cfg["aliases"]:
            pat = re.compile(r"\b" + re.escape(a) + r"\b", re.I)
            if pat.search(cleaned):
                hit = hit or cfg["label"]
                cleaned = pat.sub(" ", cleaned)
    return hit, cleaned


def best_match(toks: set[str], cands: list[Task], threshold: int):
    scored = [(len(toks & tokens(c.name)), c) for c in cands]
    scored = [(s, c) for s, c in scored if s >= threshold and s > 0]
    if not scored:
        return None, 0, 0
    top = max(s for s, _ in scored)
    tied = sorted([c for s, c in scored if s == top],
                  key=lambda c: c.end or date.min, reverse=True)
    return tied[0], top, len(tied)


def match_text(text: str, down: Task, tasks: list[Task]):
    """Return (upstream_task | None, confidence, hint_team)."""
    hint, cleaned = find_hint(text)
    if hint == down.team:
        hint = None
    toks = tokens(cleaned)
    n = len(toks)
    if n == 0:
        return None, "Unresolved", hint
    if hint:
        cands = [t for t in tasks if t.team == hint]
        up, score, ties = best_match(toks, cands, 1)
    else:
        same = [t for t in tasks if t.team == down.team and t is not down]
        up, score, ties = best_match(toks, same, min(2, n))
        if up is None:
            others = [t for t in tasks if t.team != down.team]
            up, score, ties = best_match(toks, others, 2)
    if up is None:
        return None, "Unresolved", hint
    conf = "Medium" if (score >= 2 or score == n) else "Low"
    if ties > 1:
        conf = "Low"
    return up, conf, hint


TRIGGER = re.compile(
    r"\b(?:needs?|needed|requires?|required|waiting (?:on|for)|blocked (?:by|until|on)|"
    r"assum(?:es|ing)|depends? on|dependent on)\b", re.I)


def extract_dependencies(tasks: list[Task], people: dict[str, str]) -> list[Dep]:
    local, glob, by_uid = {}, defaultdict(list), {}
    for t in tasks:
        local[(t.team, t.orig_id.lower())] = t
        glob[t.orig_id.lower()].append(t)
        by_uid[t.tid.lower()] = t
    letter_ids = sorted({k for t in tasks for k in (t.orig_id.lower(), t.tid.lower())
                         if re.search(r"[a-z]", k)}, key=len, reverse=True)
    embedded = (re.compile(r"(?<![\w.:-])(" + "|".join(re.escape(i) for i in letter_ids)
                           + r")(?![\w.:-])", re.I) if letter_ids else None)

    deps: list[Dep] = []
    by_down = defaultdict(list)
    seen = {}

    def lookup(ref: str, down: Task):
        """Find a task by ID: this plan first, then any plan if the ID is unique."""
        r = ref.strip().lower()
        t = local.get((down.team, r))
        if t is not None and t is not down:
            return t, False
        t = by_uid.get(r)
        if t is not None and t is not down:
            return t, False
        others = [x for x in glob.get(r, []) if x is not down]
        if len(others) == 1:
            return others[0], False
        return None, len(others) > 1

    def add(dep: Dep):
        if dep.up is not None:
            key = (id(dep.down), id(dep.up))
            if key in seen:
                if dep.assumed and not seen[key].assumed:
                    seen[key].assumed = dep.assumed
                return
            seen[key] = dep
        deps.append(dep)
        by_down[id(dep.down)].append(dep)

    # 1) The plan's own dependency column
    for t in tasks:
        if t.deps_raw.lower() in EMPTY_DEPS:
            continue
        for piece in re.split(r"[;,]", t.deps_raw):
            piece = piece.strip()
            if not piece or piece.lower() in EMPTY_DEPS:
                continue
            up, ambiguous = lookup(piece, t)
            if up is not None:
                add(Dep(t, up, piece, "Plan column", "High"))
                continue
            if ambiguous:
                add(Dep(t, None, piece, "Plan column", "Unresolved",
                        note="This ID exists in more than one plan; use the PREFIX:ID form"))
                continue
            refs = embedded.findall(piece) if embedded else []
            if refs:
                for ref in refs:
                    up, ambiguous = lookup(ref, t)
                    if up is not None:
                        add(Dep(t, up, piece, "Plan column", "High"))
                    elif ambiguous:
                        add(Dep(t, None, ref, "Plan column", "Unresolved",
                                note="This ID exists in more than one plan; use the PREFIX:ID form"))
                continue
            if ID_LIKE.fullmatch(piece):
                add(Dep(t, None, piece, "Plan column", "Unresolved",
                        note="ID not found in any plan"))
                continue
            if re.search(r"[A-Za-z]", piece):
                up, conf, hint = match_text(piece, t, tasks)
                add(Dep(t, up, piece, "Plan column", conf, hint))

    # 2) Dependency language hidden in notes
    for t in tasks:
        if not t.notes:
            continue
        for sent in NOTE_SPLIT.split(t.notes):
            m = TRIGGER.search(sent)
            if not m or re.search(r"needed by", sent, re.I):
                continue
            if persons_in(sent, people) - {t.owner}:
                continue                                  # a resource request, not a dependency
            remainder = sent[:m.start()] + " " + sent[m.end():]
            dates = extract_dates(remainder, t.start or date.today())
            assumed = dates[0][1] if dates else None

            hit = False                                   # explicit task IDs written in the note
            for ref in (embedded.findall(sent) if embedded else []):
                up, _ = lookup(ref, t)
                if up is not None:
                    add(Dep(t, up, sent.strip(), "Notes", "High", assumed=assumed))
                    hit = True
            if hit:
                continue

            clean = strip_spans(remainder, [d[2] for d in dates])
            hint, cleaned = find_hint(clean)
            toks = tokens(cleaned)
            if not toks:
                continue
            # Does this sentence just add an assumed date to a dependency we already have?
            best, best_overlap = None, 0
            for d in by_down[id(t)]:
                if d.up is None:
                    continue
                ov = len(toks & tokens(d.up.name))
                if ov > best_overlap:
                    best, best_overlap = d, ov
            if best is not None:
                if assumed and not best.assumed:
                    best.assumed = assumed
                continue
            # Same dependency already captured from the plan column?
            own_hint = hint if hint != t.team else None
            if hint and any(d.hint_team == own_hint
                            and len(toks & tokens(find_hint(d.raw_text)[1])) >= 1
                            for d in by_down[id(t)]):
                continue
            up, conf, h2 = match_text(clean.strip(), t, tasks)
            add(Dep(t, up, sent.strip(), "Notes", conf, h2, assumed))

    order = {t.tid: i for i, t in enumerate(tasks)}
    deps.sort(key=lambda d: (order[d.down.tid], d.up.tid if d.up else "~"))
    for i, d in enumerate(deps, 1):
        d.dep_id = f"D-{i:03d}"
    return deps


# ----------------------------------------------------------------------------
# Optional Claude pass: resolve free-text dependencies
# ----------------------------------------------------------------------------

RESOLVE_TOOL = {
    "name": "record_dependency_matches",
    "description": "For each dependency description, name the upstream task it refers to.",
    "input_schema": {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "matches": {
                "type": "array",
                "items": {
                    "type": "object",
                    "additionalProperties": False,
                    "properties": {
                        "dep_id": {"type": "string"},
                        "upstream_task_id": {
                            "type": "string",
                            "description": "A task ID from the task list, or NONE if no task fits."},
                        "confidence": {"type": "string", "enum": ["High", "Medium", "Low"]},
                        "reason": {"type": "string"}
                    },
                    "required": ["dep_id", "upstream_task_id", "confidence", "reason"]
                }
            }
        },
        "required": ["matches"]
    },
    "strict": True
}


def _shortlist(batch: list[Dep], tasks: list[Task]) -> list[Task]:
    """For big programs: candidate tasks for a batch (the hinted team's tasks plus the
    closest keyword matches), so the prompt stays a manageable size."""
    pool = {}
    for d in batch:
        hint, cleaned = find_hint(d.raw_text)
        toks = tokens(cleaned)
        ranked = sorted((t for t in tasks if t is not d.down),
                        key=lambda t: -len(toks & tokens(t.name)))
        if hint:
            for t in [t for t in ranked if t.team == hint][:150]:
                pool[t.tid] = t
        for t in ranked[:15]:
            pool[t.tid] = t
        pool[d.down.tid] = d.down
    return list(pool.values())


def resolve_with_claude(tasks: list[Task], deps: list[Dep], client=None,
                        batch_size: int = 25, full_list_limit: int = 300) -> int:
    targets = [d for d in deps if d.up is None or d.confidence == "Low"]
    if not targets:
        return 0
    if client is None:
        import anthropic
        client = anthropic.Anthropic()          # reads ANTHROPIC_API_KEY from env

    by_id = {t.tid: t for t in tasks}
    changed = 0
    for i in range(0, len(targets), batch_size):
        batch = targets[i:i + batch_size]
        pool = tasks if len(tasks) <= full_list_limit else _shortlist(batch, tasks)
        task_lines = "\n".join(
            f"{t.tid} | {t.team} | {t.name} | {t.window} | owner: {t.owner}" for t in pool)
        dep_lines = "\n".join(
            f"{d.dep_id} | dependent task {d.down.tid} ({d.down.team}): {d.down.name} | "
            f"description: \"{d.raw_text}\"" for d in batch)
        prompt = (
            "Several teams wrote separate project plans. Below is the task list, then "
            "dependency descriptions where one task needs something from another.\n"
            "For each description, pick the single task (by ID) that produces what the "
            "dependent task is waiting for. Rules: choose only IDs from the task list; answer "
            "NONE if no listed task clearly produces it (do not guess); a task cannot depend "
            "on itself.\n\n"
            f"TASKS:\n{task_lines}\n\nDEPENDENCIES TO RESOLVE:\n{dep_lines}")

        response = client.messages.create(
            model=MODEL, max_tokens=2000,
            tools=[RESOLVE_TOOL],
            tool_choice={"type": "tool", "name": "record_dependency_matches"},
            messages=[{"role": "user", "content": prompt}])

        result = None
        for block in response.content:
            if block.type == "tool_use":
                result = block.input
        if not result:
            raise RuntimeError("No tool_use block in Claude response.")

        by_dep = {d.dep_id: d for d in batch}
        for m in result["matches"]:
            d = by_dep.get(m["dep_id"])
            if d is None:
                continue
            up = by_id.get(m["upstream_task_id"])
            if up is None or up is d.down:
                d.note = f"Claude: no matching task ({m['reason']})"
                continue
            d.up, d.confidence, d.source = up, m["confidence"], "Claude"
            d.note = f"Claude: {m['reason']}"
            changed += 1

    # drop duplicates created by re-pointing
    seen, keep = set(), []
    for d in deps:
        key = (d.down.tid, d.up.tid) if d.up else None
        if key and key in seen:
            continue
        if key:
            seen.add(key)
        keep.append(d)
    deps[:] = keep
    return changed


# ----------------------------------------------------------------------------
# Timeline conflicts
# ----------------------------------------------------------------------------


def sev_by_days(days: int, high: int, med: int) -> str:
    return "High" if days >= high else "Medium" if days >= med else "Low"


def timeline_conflicts(tasks: list[Task], deps: list[Dep]) -> list[dict]:
    out = []

    def add(**kw):
        kw.setdefault("dep_id", "")
        kw.setdefault("confidence", "High")
        kw.setdefault("date_a", None)
        kw.setdefault("date_b", None)
        kw.setdefault("days", None)
        out.append(kw)

    # --- A. Dependency timing + assumptions -------------------------------
    for d in deps:
        U, D = d.up, d.down
        if U is None or not (U.start and U.end and D.start and D.end):
            continue
        teams = f"{U.team} -> {D.team}" if U.team != D.team else U.team
        overlap = (U.end - D.start).days
        sev = desc = None
        if D.is_milestone or U.is_milestone:
            if D.start < U.end:
                sev, desc = "High", (
                    f"{D.tid} '{D.name}' is scheduled for {D.start:%b %d}, before its "
                    f"prerequisite {U.tid} '{U.name}' finishes ({U.end:%b %d}).")
        elif D.end <= U.end:
            sev = "High"
            when = ("the same day" if D.end == U.end
                    else f"{(U.end - D.end).days} day(s) after")
            desc = (f"{D.tid} '{D.name}' ({D.window}) finishes {when} its prerequisite "
                    f"{U.tid} '{U.name}' finishes ({U.end:%b %d}). The dependency "
                    f"cannot be met as planned.")
        elif D.start < U.end:
            sev = "Medium" if overlap >= 3 else "Low"
            desc = (f"{D.tid} '{D.name}' starts {D.start:%b %d}, {overlap} day(s) before "
                    f"prerequisite {U.tid} '{U.name}' finishes ({U.end:%b %d}).")
        elif D.start == U.end:
            sev, desc = "Low", (
                f"{D.tid} starts the same day prerequisite {U.tid} finishes "
                f"({U.end:%b %d}); no buffer.")
        if sev:
            add(type="Dependency timing", severity=sev, teams=teams,
                tasks=f"{U.tid}, {D.tid}", description=desc, date_a=U.end, date_b=D.start,
                days=overlap, dep_id=d.dep_id, confidence=d.confidence,
                action=(f"Agree a firm handoff date between {U.team} and {D.team}, or "
                        f"resequence {D.tid}. Confirm whether {D.tid} can start on partial "
                        f"inputs."))
        if d.assumed and U.end > d.assumed:
            late = (U.end - d.assumed).days
            add(type="Assumption vs upstream plan", severity="High" if late >= 5 else "Medium",
                teams=teams, tasks=f"{U.tid}, {D.tid}",
                description=(f"{D.team} assumes '{U.name}' ({U.tid}) is ready by "
                             f"{d.assumed:%b %d}, but {U.team}'s plan shows it finishing "
                             f"{U.end:%b %d} ({late} days later)."),
                date_a=d.assumed, date_b=U.end, days=late, dep_id=d.dep_id,
                confidence=d.confidence,
                action=f"Align {D.team}'s assumption with {U.team}'s committed date, or "
                       f"pull {U.tid} in.")

    # --- B. Milestone date mismatch ---------------------------------------
    people = build_people(tasks)
    for word in MILESTONE_WORDS:
        claims = []                                   # (date, task)
        for t in tasks:
            if not t.notes or word not in t.notes.lower():
                continue
            for sent in NOTE_SPLIT.split(t.notes):
                if word in sent.lower():
                    ds = extract_dates(sent, t.start or date.today())
                    if ds:
                        claims.append((ds[0][0] if False else ds[0][1], t))
        distinct = sorted({c[0] for c in claims})
        if len(distinct) > 1:
            by_date = defaultdict(list)
            for dt, t in claims:
                by_date[dt].append(t)
            detail = "; ".join(
                f"{dt:%b %d}: " + ", ".join(sorted({f'{t.team} ({t.tid})' for t in ts}))
                for dt, ts in sorted(by_date.items()))
            add(type="Milestone date mismatch", severity="High",
                teams=", ".join(sorted({t.team for _, t in claims})),
                tasks=", ".join(sorted({t.tid for _, t in claims})),
                description=(f"Teams disagree on the {word} date. {detail}."),
                date_a=distinct[0], date_b=distinct[-1],
                days=(distinct[-1] - distinct[0]).days,
                action=f"Hold a single decision on the {word} date and update every "
                       f"plan to match.")

    # --- C. Gates that finish after cutover -------------------------------
    cut = [t for t in tasks if t.is_core and "cutover" in t.name.lower() and t.start]
    if cut:
        cutover = min(cut, key=lambda t: t.start)
        for t in tasks:
            if t is cutover or not t.end or not GATE_PATTERN.search(t.name):
                continue
            if t.end > cutover.start:
                linked = any(d.up is t and d.down.is_core for d in deps)
                add(type="Gate after cutover", severity="High", teams=t.team,
                    tasks=f"{t.tid}, {cutover.tid}",
                    description=(f"'{t.name}' ({t.tid}, {t.team}) finishes {t.end:%b %d}, "
                                 f"{(t.end - cutover.start).days} day(s) after the Core cutover "
                                 f"({cutover.start:%b %d}). "
                                 + ("" if linked else "No Core task depends on it, so "
                                    "cutover could proceed without it.")),
                    date_a=cutover.start, date_b=t.end,
                    days=(t.end - cutover.start).days,
                    action=f"Make {t.tid} a formal gate in the Core plan (before Go/No-Go) "
                           f"and pull it earlier, or move the cutover.")

    # --- D. Holidays ------------------------------------------------------
    for t in tasks:
        if not t.start or not t.end:
            continue
        wd = list(weekdays_between(t.start, t.end))
        lost = [d for d in wd if d in HOLIDAYS]
        if lost:
            share = len(lost) / len(wd)
            names = ", ".join(f"{d:%b %d} {HOLIDAYS[d]}" for d in lost)
            add(type="Holiday impact", severity="Medium" if share >= 0.25 else "Low",
                teams=t.team, tasks=t.tid,
                description=(f"'{t.name}' ({t.window}) includes {names}: "
                             f"{len(wd) - len(lost)} usable working days of {len(wd)}."),
                date_a=t.start, date_b=t.end, days=len(lost),
                action="Extend the task or confirm the team will work through the holiday.")

    # --- E. Change freezes / blackouts stated in notes --------------------
    for t in tasks:
        if not t.notes or not t.start or not t.end:
            continue
        for sent in re.split(r"[;]\s*", t.notes):
            if not re.search(r"freeze|blackout", sent, re.I):
                continue
            ds = extract_dates(sent, t.start)
            if len(ds) < 2:
                continue
            f_start, f_end = ds[0][0], ds[1][1]
            lo, hi = max(f_start, t.start), min(f_end, t.end)
            if lo <= hi:
                overlap_days = (hi - lo).days + 1
                total = (t.end - t.start).days + 1
                add(type="Change freeze overlap",
                    severity="High" if overlap_days / total >= 0.5 else "Medium",
                    teams=t.team, tasks=t.tid,
                    description=(f"'{t.name}' ({t.window}) overlaps a freeze/blackout "
                                 f"{f_start:%b %d} - {f_end:%b %d} for {overlap_days} of "
                                 f"{total} days. Note says: \"{sent.strip()}\"."),
                    date_a=f_start, date_b=t.end, days=overlap_days,
                    action="Move the task out of the freeze window or arrange explicit "
                           "coverage/exception.")

    out.sort(key=lambda c: (SEV_ORDER[c["severity"]], c["date_b"] or date.max))
    for i, c in enumerate(out, 1):
        c["id"] = f"TC-{i:03d}"
    return out


# ----------------------------------------------------------------------------
# Resource conflicts
# ----------------------------------------------------------------------------


def resource_conflicts(tasks: list[Task]) -> list[dict]:
    people = build_people(tasks)
    commits = defaultdict(list)        # person -> [dict(task, role, alloc, start, end)]
    for t in tasks:
        if t.owner in people.values() and t.start and t.end:
            commits[t.owner].append(dict(task=t, role="Owner", alloc=None,
                                         start=t.start, end=t.end))
        if not t.notes or not t.start or not t.end:
            continue
        for p in persons_in(t.notes, people) - {t.owner}:
            sent = next((s for s in NOTE_SPLIT.split(t.notes)
                         if p.lower() in s.lower() or p.split()[0].lower() in s.lower()),
                        t.notes)
            pct = re.search(r"(\d{1,3})\s*%", sent)
            ds = extract_dates(sent, t.start)
            s, e = (ds[0][0], ds[0][1]) if ds else (t.start, t.end)
            commits[p].append(dict(task=t, role="Support requested",
                                   alloc=int(pct.group(1)) if pct else None,
                                   start=s, end=e))

    home = {}
    for t in tasks:
        if t.owner:
            home.setdefault(t.owner, t.team)

    out = []
    for person, items in commits.items():
        lo = min(i["start"] for i in items)
        hi = max(i["end"] for i in items)
        segments = []                   # [start, end, frozenset(index)]
        for d in weekdays_between(lo, hi):
            active = frozenset(n for n, i in enumerate(items) if i["start"] <= d <= i["end"])
            if segments and segments[-1][2] == active and \
                    (d - segments[-1][1]).days <= 3:
                segments[-1][1] = d
            else:
                segments.append([d, d, active])
        for s, e, active in segments:
            alloc_sum = sum(items[n]["alloc"] or 0 for n in active)
            if len(active) >= RESOURCE_THRESHOLD or alloc_sum > 100:
                detail = "; ".join(
                    f"{items[n]['task'].tid} {items[n]['task'].name} "
                    f"[{items[n]['task'].team}; {items[n]['role']}"
                    + (f", {items[n]['alloc']}%" if items[n]["alloc"] else "") + "]"
                    for n in sorted(active))
                sev = "High" if len(active) >= RESOURCE_THRESHOLD + 1 or alloc_sum > 100 \
                    else "Medium"
                teams = sorted({items[n]["task"].team for n in active})
                out.append(dict(
                    person=person, home=home.get(person, ""), start=s, end=e,
                    ids=[items[n]["task"].tid for n in sorted(active)],
                    count=len(active), alloc=alloc_sum or None, detail=detail,
                    teams=", ".join(teams), severity=sev,
                    action=(f"Confirm {person}'s real availability for {s:%b %d}-{e:%b %d}. "
                            f"Stagger the requests, name a backup, or have one program own "
                            f"the priority call.")))
    out.sort(key=lambda r: (SEV_ORDER[r["severity"]], r["start"]))
    for i, r in enumerate(out, 1):
        r["id"] = f"RC-{i:03d}"
    return out


# ----------------------------------------------------------------------------
# Risks
# ----------------------------------------------------------------------------


def categorize(text: str) -> str:
    t = text.lower()
    if re.search(r"freeze|holiday|reduced|blackout", t):
        return "Calendar"
    if re.search(r"waiting|blocked|block\b|depend|needs", t):
        return "Dependency"
    if re.search(r"timeout|critical|finding|bug|defect", t):
        return "Quality"
    if re.search(r"hardware|lead time|capacity", t):
        return "Supply / capacity"
    return "Schedule"


def build_risks(tasks, deps, tc, rc) -> list[dict]:
    risks = []

    def add(**kw):
        kw.setdefault("conflicts", "")
        kw.setdefault("status", "")
        risks.append(kw)

    # 1) Reported status and risk-flavored notes, one risk per task
    for t in tasks:
        parts, src = [], []
        if t.status and BAD_STATUS.search(t.status):
            parts.append(f"Reported status: {t.status}.")
            src.append("Reported status")
        hits = [s.strip() for s in NOTE_SPLIT.split(t.notes) if RISK_WORDS.search(s)]
        if hits:
            parts.append("Notes: " + "; ".join(hits) + ".")
            src.append("Notes")
        if parts:
            sev = "High" if (t.status and re.search(r"red|blocked", t.status, re.I)) else "Medium"
            add(category=categorize(" ".join(parts)), severity=sev, team=t.team,
                tasks=t.tid, description=f"{t.name} ({t.tid}): " + " ".join(parts),
                source=" + ".join(src), status=t.status,
                mitigation="Assign an owner and a date to resolve; report it in the "
                           "exec risk log.")

    # 2) Dependencies that are unresolved or only weakly matched
    for d in deps:
        if d.up is None:
            if d.source == "Notes":
                continue            # already covered by the note-based risk for this task
            where = d.hint_team or "any plan"
            add(category="Dependency", severity="Medium", team=d.down.team,
                tasks=d.down.tid,
                description=(f"{d.down.tid} '{d.down.name}' depends on \"{d.raw_text}\", "
                             f"but no matching task exists on {where}. Nobody visibly owns "
                             f"the work."),
                source="Dependency gap", status=d.down.status,
                mitigation=f"Ask {where} to add an owned task with a date, or confirm the "
                           f"dependency is not real.")
        elif d.confidence == "Low" and d.up.team != d.down.team:
            add(category="Dependency", severity="Low", team=d.down.team,
                tasks=f"{d.down.tid}, {d.up.tid}",
                description=(f"{d.down.tid} depends on \"{d.raw_text}\"; closest match is "
                             f"{d.up.tid} '{d.up.name}' ({d.up.team}) but the match is "
                             f"weak."),
                source="Dependency gap", status=d.down.status,
                mitigation=f"Confirm with {d.up.team} that {d.up.tid} delivers this, or "
                           f"add the missing deliverable to their plan.")

    # 3) High-severity conflicts become risks
    for c in tc:
        if c["severity"] == "High":
            add(category="Calendar" if c["type"] in ("Holiday impact", "Change freeze overlap")
                else "Schedule",
                severity="High", team=c["teams"], tasks=c["tasks"],
                description=f"{c['type']}: {c['description']}",
                source=f"Derived ({c['id']})", conflicts=c["id"],
                mitigation=c["action"])
    for r in rc:
        if r["severity"] == "High":
            add(category="Resource", severity="High", team=r["teams"],
                tasks=", ".join(r["ids"]),
                description=(f"{r['person']} is committed to {r['count']} things at once "
                             f"({r['start']:%b %d} - {r['end']:%b %d})."),
                source=f"Derived ({r['id']})", conflicts=r["id"], mitigation=r["action"])

    # 4) Tasks reported Green/On Track but caught in a High conflict
    by_id = {t.tid: t for t in tasks}
    flagged = defaultdict(list)
    for c in tc:
        if c["severity"] != "High":
            continue
        for tid in c["tasks"].split(", "):
            t = by_id.get(tid)
            if t and GOOD_STATUS.match(t.status):
                flagged[tid].append(c["id"])
    for tid, ids in flagged.items():
        t = by_id[tid]
        add(category="Reporting", severity="Medium", team=t.team, tasks=tid,
            description=(f"{tid} '{t.name}' is reported '{t.status}' but appears in "
                         f"{len(ids)} high-severity conflict(s): {', '.join(ids)}."),
            source="Status mismatch", status=t.status, conflicts=", ".join(ids),
            mitigation="Re-assess the status with the owner; Green is not credible "
                       "while the conflict stands.")

    risks.sort(key=lambda r: (SEV_ORDER[r["severity"]], r["team"]))
    for i, r in enumerate(risks, 1):
        r["id"] = f"R-{i:03d}"
    return risks


# ----------------------------------------------------------------------------
# Excel output
# ----------------------------------------------------------------------------

FONT = "Arial"
SEV_FILL = {"High": "F8CBAD", "Medium": "FFE699", "Low": "C6E0B4"}


# Backslashes make this a custom format, so Excel shows mm-dd-yy exactly
# (a bare "mm-dd-yy" is mapped by Excel to your regional short-date setting).
DATE_FORMAT = "mm\\-dd\\-yy"
TRACKING_HEADERS = ["Action Taken", "Owner", "Closed Date"]
TRACKING_WIDTHS = [40, 18, 13]


def write_sheet(wb, title, headers, rows, widths, date_cols=(), sev_col=None):
    n_data_cols = len(headers)
    headers = list(headers) + TRACKING_HEADERS          # blank columns for you to fill in
    widths = list(widths) + TRACKING_WIDTHS
    ws = wb.create_sheet(title)
    ws.append(headers)
    for r in rows:
        ws.append(list(r) + [None, None, None])
    for c in ws[1]:
        c.font = Font(name=FONT, bold=True, color="FFFFFF", size=10)
        c.fill = PatternFill("solid", fgColor="1F3864" if c.column <= n_data_cols else "385723")
        c.alignment = Alignment(vertical="center", wrap_text=True)
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.font = Font(name=FONT, size=10)
            c.alignment = Alignment(vertical="top", wrap_text=True)
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for col in tuple(date_cols) + (n_data_cols + 3,):    # last tracking column = Closed Date
        for r in range(2, ws.max_row + 1):
            ws.cell(r, col).number_format = DATE_FORMAT
    for r in range(2, ws.max_row + 1):                   # light yellow = cells to fill in
        for col in range(n_data_cols + 1, n_data_cols + 4):
            ws.cell(r, col).fill = PatternFill("solid", fgColor="FFF9E5")
    if sev_col:
        for r in range(2, ws.max_row + 1):
            v = ws.cell(r, sev_col).value
            if v in SEV_FILL:
                ws.cell(r, sev_col).fill = PatternFill("solid", fgColor=SEV_FILL[v])
                ws.cell(r, sev_col).font = Font(name=FONT, size=10, bold=True)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    return ws


def write_workbook(path, deps, tc, rc, risks):
    wb = Workbook()
    wb.remove(wb.active)

    conflict_by_dep = defaultdict(list)
    for c in tc:
        if c["dep_id"]:
            conflict_by_dep[c["dep_id"]].append(c["id"])

    waiting = defaultdict(set)         # upstream task ID -> downstream task IDs waiting on it
    for d in deps:
        if d.up is not None:
            waiting[d.up.tid].add(d.down.tid)

    dep_rows = []
    for d in sorted(deps, key=lambda d: d.dep_id):
        U, D = d.up, d.down
        if d.up is None:
            status = "Unresolved"
        elif conflict_by_dep.get(d.dep_id):
            status = "Conflict"
        elif d.confidence == "Low":
            status = "Unconfirmed match"
        else:
            status = "OK"
        float_days = (D.start - U.end).days if (U and U.end and D.start) else None
        dep_rows.append([
            d.dep_id, D.team, D.tid, D.name, D.start, D.end,
            U.team if U else (d.hint_team or ""), U.tid if U else "", U.name if U else "",
            U.start if U else None, U.end if U else None,
            len(waiting[U.tid]) if U else None, d.kind, d.source, d.confidence,
            d.raw_text, d.assumed, float_days, status,
            ", ".join(conflict_by_dep.get(d.dep_id, [])), d.note])
    write_sheet(wb, "Dependencies", [
        "Dep ID", "Downstream (Waiting) Team", "Downstream (Waiting) Task ID",
        "Downstream (Waiting) Task", "Downstream Start", "Downstream End",
        "Upstream (Providing) Team", "Upstream (Providing) Task ID",
        "Upstream (Providing) Task", "Upstream Start", "Upstream End",
        "# Tasks Waiting on This Upstream Task", "Type", "Found In", "Match Confidence", "Original Text",
        "Assumed Ready Date", "Float (days, negative = overlap)", "Status",
        "Related Conflicts", "Match Notes"],
        dep_rows, [9, 22, 13, 34, 12, 12, 22, 13, 34, 12, 12, 14, 13, 13, 12, 40, 12, 14, 16, 14, 34],
        date_cols=(5, 6, 10, 11, 17))

    write_sheet(wb, "Risks", [
        "Risk ID", "Severity", "Category", "Team(s)", "Related Tasks", "Description",
        "Source", "Reported Status", "Related Conflicts", "Suggested Mitigation"],
        [[r["id"], r["severity"], r["category"], r["team"], r["tasks"], r["description"],
          r["source"], r["status"], r["conflicts"], r["mitigation"]] for r in risks],
        [8, 10, 16, 24, 16, 70, 22, 14, 14, 50], sev_col=2)

    write_sheet(wb, "Timeline Conflicts", [
        "Conflict ID", "Severity", "Type", "Team(s)", "Related Tasks", "Description",
        "Date A", "Date B", "Days", "Related Dep ID", "Match Confidence", "Suggested Action"],
        [[c["id"], c["severity"], c["type"], c["teams"], c["tasks"], c["description"],
          c["date_a"], c["date_b"], c["days"], c["dep_id"], c["confidence"], c["action"]]
         for c in tc],
        [10, 10, 24, 28, 14, 75, 12, 12, 7, 11, 12, 50], date_cols=(7, 8), sev_col=2)

    write_sheet(wb, "Resource Conflicts", [
        "Conflict ID", "Severity", "Person", "Home Team", "Window Start", "Window End",
        "Concurrent Commitments", "Stated Allocation (%)", "Teams Involved",
        "Commitments", "Suggested Action"],
        [[r["id"], r["severity"], r["person"], r["home"], r["start"], r["end"], r["count"],
          r["alloc"], r["teams"], r["detail"], r["action"]] for r in rc],
        [10, 10, 16, 20, 12, 12, 13, 13, 40, 80, 50], date_cols=(5, 6), sev_col=2)

    try:
        wb.save(path)
    except PermissionError:
        raise SystemExit(f"ERROR: cannot write {path}. If it is open in Excel, close it "
                         f"and run again.")


# ----------------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------------


def collect_files(inputs: list[str], exclude: list[str] = ()) -> list[str]:
    files = []
    for p in inputs:
        if os.path.isdir(p):
            files += sorted(glob.glob(os.path.join(p, "*.xlsx")))
        elif os.path.isfile(p):
            if not p.lower().endswith(".xlsx"):
                raise SystemExit(f"ERROR: '{p}' is not an .xlsx file.")
            files.append(p)
        else:
            raise SystemExit(
                f"ERROR: '{p}' was not found.\n"
                f"The script is looking from this folder: {os.getcwd()}\n"
                f"Pass the full path to the folder holding the plan .xlsx files, "
                f"or cd into the folder that contains it.")
    skip = {os.path.abspath(e) for e in exclude}
    return [f for f in files
            if not os.path.basename(f).startswith("~$") and os.path.abspath(f) not in skip]


def main():
    ap = argparse.ArgumentParser(description="Synthesize multiple team project plans.")
    ap.add_argument("inputs", nargs="+", help="plan .xlsx files, or a folder of them")
    ap.add_argument("-o", "--output", default="plan_review.xlsx")
    ap.add_argument("--use-claude", action="store_true",
                    help="use Claude to resolve free-text dependencies the rules can't match")
    ap.add_argument("--core", default=None,
                    help="word in the core plan's file name (marks the core team for "
                         "cutover/gate checks)")
    ap.add_argument("--default-year", type=int, default=date.today().year,
                    help="year for dates written without one (auto-detected from other dates)")
    args = ap.parse_args()

    files = collect_files(args.inputs, exclude=[args.output])
    if not files:
        raise SystemExit("No .xlsx plan files found in the given path(s).")
    cfgs = resolve_teams(files, args.core)
    configure_teams(cfgs)
    year = infer_year(files, args.default_year)
    print(f"Plans found: {len(files)}   year used for dates without one: {year}")

    warnings: list[str] = []
    tasks: list[Task] = []
    for f, cfg in zip(files, cfgs):
        tasks += load_plan(f, cfg, year, warnings)
    uniquify_ids(tasks, warnings)
    finalize_aliases(cfgs, tasks)
    print(f"\nTeams: {len({t.team for t in tasks})}   Tasks: {len(tasks)}")
    if not any(t.is_core for t in tasks):
        warnings.append("No core team identified, so the cutover/gate-after-cutover checks "
                        "were skipped. Use --core <word in the core plan's file name>.")

    people = build_people(tasks)
    deps = extract_dependencies(tasks, people)

    if args.use_claude:
        try:
            n = resolve_with_claude(tasks, deps)
            print(f"\nClaude resolved {n} dependency description(s).")
            for i, d in enumerate(deps, 1):
                d.dep_id = f"D-{i:03d}"
        except Exception as e:                       # keep the rule-based result
            warnings.append(f"Claude pass failed ({e}); using rule-based matches only.")

    tc = timeline_conflicts(tasks, deps)
    rc = resource_conflicts(tasks)
    risks = build_risks(tasks, deps, tc, rc)
    write_workbook(args.output, deps, tc, rc, risks)

    cross = sum(1 for d in deps if d.kind == "Cross-team")
    unres = sum(1 for d in deps if d.kind == "Unresolved")
    print("\n=== SUMMARY ===")
    print(f"Tasks: {len(tasks)}  |  Dependencies: {len(deps)} "
          f"({cross} cross-team, {unres} unresolved)")
    for name, rows in (("Timeline conflicts", tc), ("Resource conflicts", rc), ("Risks", risks)):
        c = Counter(r["severity"] for r in rows)
        print(f"{name}: {len(rows)}  (High {c['High']}, Medium {c['Medium']}, Low {c['Low']})")
    for w in warnings:
        print("WARNING:", w)
    print(f"\nSaved {args.output}")


if __name__ == "__main__":
    main()
