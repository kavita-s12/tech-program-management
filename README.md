# Tech Program Management

Tools I've built with Claude to take the manual, repetitive work out of Technical Program Management: turning raw project information into clear, trackable follow-ups, and surfacing problems early. Each project is a self-contained folder with its own README, sample data, and setup steps.

## Projects

### [Meeting Notes: Decision, Action, Risk Tracker](meeting-notes-decision-action-risk-tracker/)

Turns meeting transcripts into structured decisions, owned action items with due dates, and open risks, using the Claude API. Results accumulate in a formatted Excel tracker across meetings, with duplicate detection so items that come up again in different words are not logged twice. Available as a Streamlit app and a command-line script, with a human review step before anything is shared.

### [Multi-Team Project Plan Synthesizer](multi-team-project-plan-synthesizer/)

Combines the separate project plans of a core team, a vendor, and any number of partner teams into one review workbook covering cross-team dependencies, timeline conflicts, resource conflicts, and risks. Rule-based checks do the date math, with an optional Claude pass to resolve dependencies written in plain English. Each finding comes with columns for tracking the action taken, owner, and closed date, so teams can fix their project plans and re-run the check.

## Why these projects

In technical delivery and process optimization initiatives, the primary challenge is rarely a single project plan or meeting. Instead, issues usually stem from gaps between cross-functional teams due to missed dependencies, conflicting priorities, and mismatched stakeholder plans. These tools identify those gaps efficiently, offering two key benefits:

1. Save time for more productive work.
2. Enable teams to proactively manage risks and address gaps before they impact delivery on the critical path.

## Tech

Python, Claude API (Anthropic), Streamlit, openpyxl
