# Multi-Team Project Plan Synthesizer

A tool that reads separate project plans from a core team, a vendor, and partner teams, and turns them into one review workbook: cross-team **dependencies**, **timeline conflicts**, **resource conflicts**, and **risks**, each ready to be worked through and tracked to closure.

## Why I built this

For Technical Delivery initiatives, there are multiple teams involved, each with its own set of priorities and often its own project plan. Each plan looks reasonable on its own. However, an integrated view of the project plan for the initiative is critical to the success of the program. It's also where the problems show up in the gaps between them: a vendor assumes an environment is ready a week before the core team will deliver it, two teams quietly plan for different cutover dates, one engineer is requested by three teams in the same fortnight, or a compliance sign-off lands after go-live. Finding these by hand means lining up several spreadsheets in different formats, which is slow and easy to get wrong.

This tool automates that first pass. It does the date math with plain, repeatable rules, uses Claude only where language understanding is needed, and leaves the decisions to people.

## What it does

1. **Input**: a folder of `.xlsx` plans, one per team. There is no limit on the number of plans, and each team's name is picked up from its file name. The plans can use different column names, date formats, and dependency styles.
2. **Normalize**: every plan is mapped into one common task list (team, task ID, owner, start, end, dependencies, status, notes).
3. **Extract and check**:
   - **Dependencies**: from each plan's own dependency column, and from dependency language buried in notes ("Needs final rate-limit values from Atlas core team"). Each dependency is marked within-team, cross-team, or unresolved, with a match confidence.
   - **Timeline conflicts**: downstream tasks scheduled before their prerequisite finishes, teams assuming a ready date their upstream plan doesn't meet, teams disagreeing on a milestone date (such as cutover), sign-offs or certifications that finish after cutover, holidays inside a task window, and change freezes that overlap a task.
   - **Resource conflicts**: a person committed to too many things at once, including support requests written in other teams' notes.
   - **Risks**: from reported status (Amber, At Risk), risk wording in notes, dependencies nobody visibly owns, high-severity conflicts, and tasks reported Green that sit inside a high-severity conflict.
4. **Output**: one Excel workbook with four tabs (Dependencies, Risks, Timeline Conflicts, Resource Conflicts). Findings are sorted and color-coded by severity, and every tab ends with blank **Action Taken**, **Owner**, and **Closed Date** columns for tracking the resolution.
5. **Iterate**: teams fix their plans and resubmit. Re-running the tool on the revised plans checks that resolved conflicts are gone and that the fixes didn't create new ones.

## Tech stack

- **Python**
- **openpyxl**: reading the team plans and writing the review workbook
- **Claude API** (Anthropic): optional pass that resolves free-text dependencies the rules cannot match (tool use with a strict JSON schema)

## Key design decisions

- **Rules for the math, Claude for the language.** Date comparisons, overlaps, and counts are deterministic, so the same plans always give the same findings. Claude is used only to decide which task a phrase like "Atlas staging stable" refers to, and it is optional (`--use-claude`).
- **Confidence on every match.** Dependencies written as task IDs are High. Text matches are Medium or Low. Anything the tool cannot match is reported as Unresolved instead of guessed, because an unowned dependency is itself a risk worth surfacing.
- **Messy inputs on purpose.** The sample plans use four different layouts: ISO dates, `dd-mmm-yyyy`, start plus duration with no end date, and dates with no year. One plan has no task IDs at all. This is closer to real life than a clean template.
- **Strict tool-use schema for the Claude pass.** The response must match a fixed schema, and the tool is told to answer `NONE` rather than guess. Any task ID it returns is checked against the real task list. If the call fails, the tool keeps the rule-based result and prints a warning.
- **Human in the loop.** The tool finds and ranks issues; people decide what to change. The tracking columns support that, and each round starts fresh from the revised plans.
- **Answer key for testing.** The sample plans contain deliberately seeded dependencies, conflicts, and risks, listed in `sample_plans/answer_key.md`, so the output can be checked against known answers.

## Running it locally

```bash
git clone https://github.com/kavita-s12/tech-program-management.git
cd tech-program-management/multi-team-project-plan-synthesizer
pip install -r requirements.txt
python plan_synthesizer.py sample_plans -o plan_review.xlsx
```

On Windows PowerShell the commands are the same. To use the optional Claude pass:

```powershell
$env:ANTHROPIC_API_KEY="your_key_here"
python plan_synthesizer.py sample_plans -o plan_review.xlsx --use-claude
```

macOS/Linux: `export ANTHROPIC_API_KEY=your_key_here`

Notes:
- Put as many plan files as you need in the folder. Each file is one team.
- Give each plan a distinct file name that identifies its team (for example `infra_services_plan.xlsx`). Files whose names resolve to the same team name are treated as one team.
- The cutover and "sign-off after cutover" checks need to know which plan is the core team's. Add `--core <word in the core plan's file name>` (for example `--core core`) if the tool prints a warning that no core team was identified.
- Close the plan files and any earlier review file in Excel before running. Excel locks open files, and the script will stop with a message if it can't read or write one.
- The tool reads `.xlsx` only. Save older `.xls` plans as `.xlsx` first.
- To keep each round's results, give each run a different output name (`plan_review_round1.xlsx`, `plan_review_round2.xlsx`).

## Configuration

Settings are at the top of `plan_synthesizer.py`:

| Setting | What it controls |
|---|---|
| `TEAM_CONFIG` | **Optional.** A friendly name, ID prefix, and aliases for teams you want to control (for example "Atlas" means Core). Any plan that doesn't match an entry is handled automatically: the team name comes from the file name (`5_infra_services_plan.xlsx` becomes "Infra Services") and distinctive words in that name work as aliases in dependency text. |
| `COLUMN_ALIASES` | The header names each plan may use for ID, task, owner, start, end, duration, dependencies, status, and notes. |
| `HOLIDAYS` | Dates counted as non-working days. Add your own company holidays. |
| `RESOURCE_THRESHOLD` | How many concurrent commitments count as a resource conflict (default 3). |
| `MILESTONE_WORDS`, `GATE_PATTERN` | Words used to find milestone dates (cutover, go-live) and gate tasks (sign-off, certification). |

## Project structure

```
.
├── plan_synthesizer.py                       # Reads plans, finds issues, writes the workbook
├── requirements.txt
├── sample_plans/
│   ├── 1_core_team_plan.xlsx
│   ├── 2_vendor_team_plan.xlsx
│   ├── 3_partner_a_payments_integration_plan.xlsx
│   ├── 4_partner_b_data_security_plan.xlsx
│   └── answer_key.md                         # Issues seeded in the sample plans
└── README.md
```

## Sample results

On the four sample plans (38 tasks), the rule-based run produces 41 dependencies (9 cross-team, 3 unresolved), 20 timeline conflicts (12 High, 5 Medium, 3 Low), 1 resource conflict, and 29 risks. Among them are the planted problems: the Dec 7 vs Dec 14 cutover disagreement, three sign-offs that finish after cutover, a vendor assuming the cluster is ready a week early, and one engineer requested by four commitments in the same week. The answer key lists what the rules alone find and what needs the Claude pass.

The tool is built for large programs. A test run on 15 plans with 3,000 tasks took about 17 seconds. With `--use-claude`, dependencies are sent in batches of 25 and large task lists are shortlisted, so each prompt stays small.

## Known limitations

- Free-text dependency matching is keyword-based. It is wrong or silent when two plans use very different words for the same thing. That is the gap the optional Claude pass is meant to fill.
- All dependencies are treated as finish-to-start. A task that legitimately starts in parallel with its prerequisite will be flagged and needs a human call.
- Resource checks count concurrent commitments and stated percentages. They do not know real capacity, leave, or part-time schedules.
- The holiday list is maintained by hand.
- Each plan must have its column headers in row 1 of the first sheet. A title row above the headers, or a plan on a second tab, will not be read.
- If two plans reuse the same task ID, those tasks are shown as `PREFIX:ID` (for example `IS:T-01`) so findings stay unambiguous. A dependency is looked up in the task's own plan first.
- Task IDs written in notes are picked up when the sentence contains words like "needs" or "waiting on". Otherwise, put dependencies in the dependency column.
- Dependency, conflict, and risk IDs (D-001, TC-001, and so on) are renumbered on every run, so refer to findings by the task IDs involved when comparing rounds.

## Possible extensions

- Automatic round-over-round comparison showing resolved, new, and unchanged findings
- A combined task view (all teams on one timeline) and critical-path highlighting
- Import from Smartsheet, Jira, or MS Project exports instead of spreadsheets
- Allocation by percentage and real capacity for resource checks
- Post the high-severity findings to Slack, as in my [meeting-notes tool](../meeting-notes-decision-action-risk-tracker/)
- Carry the tracking columns forward between rounds
