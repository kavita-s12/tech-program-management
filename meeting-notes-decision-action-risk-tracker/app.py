"""
app.py

Streamlit UI for the meeting notes -> action items pipeline.

Setup:
    pip install -r requirements.txt
    $env:ANTHROPIC_API_KEY="your_key_here"      (PowerShell)
    export ANTHROPIC_API_KEY=your_key_here      (Mac/Linux)

Run:
    streamlit run app.py

The extraction schema and Excel logic live in meeting_extraction.py, shared
with extract_action_items.py (the CLI script), so both stay in sync.
"""

import json
import streamlit as st
import anthropic

from meeting_extraction import extract_from_transcript, build_workbook, EXCEL_PATH


@st.cache_resource
def get_client():
    return anthropic.Anthropic()


def post_to_slack(webhook_url: str, result: dict) -> None:
    import requests

    lines = [f"*Meeting Summary*\n{result['summary']}", ""]

    if result["decisions"]:
        lines.append("*Decisions*")
        for d in result["decisions"]:
            lines.append(f"- {d['decision']}")
        lines.append("")

    if result["action_items"]:
        lines.append("*Action Items*")
        for a in result["action_items"]:
            lines.append(f"- [{a['priority'].upper()}] {a['task']} — {a['owner']} (due {a['due_date']})")
        lines.append("")

    if result["open_risks"]:
        lines.append("*Open Risks*")
        for r in result["open_risks"]:
            line = f"- {r['risk']}"
            if r.get("mitigation_plan") and r["mitigation_plan"] != "Not specified":
                line += f" (mitigation: {r['mitigation_plan']})"
            lines.append(line)

    payload = {"text": "\n".join(lines)}
    resp = requests.post(webhook_url, json=payload, timeout=10)
    resp.raise_for_status()


# ---------- UI ----------

st.set_page_config(page_title="Meeting Notes -> Decisions, Action Items & Risks Tracker", layout="wide")
st.title("Meeting Notes -> Decisions, Action Items & Risks Tracker")
st.caption(
    "Paste a transcript, extract structured decisions, action items, and risks "
    "with Claude, review, and save to your Excel tracker, Slack, or JSON."
)

if "result" not in st.session_state:
    st.session_state.result = None

transcript = st.text_area(
    "Paste meeting transcript",
    height=300,
    placeholder="Paste your transcript here, or upload a .txt file below..."
)

uploaded = st.file_uploader("...or upload a .txt file", type=["txt"])
if uploaded is not None:
    transcript = uploaded.read().decode("utf-8")
    st.info(f"Loaded {uploaded.name} ({len(transcript)} characters)")

if st.button("Extract", type="primary", disabled=not transcript):
    with st.spinner("Calling Claude..."):
        try:
            st.session_state.result = extract_from_transcript(transcript, client=get_client())
        except Exception as e:
            st.error(f"Extraction failed: {e}")

# ---------- Review / edit extracted results ----------

if st.session_state.result:
    result = st.session_state.result
    st.divider()
    st.subheader("Review before sending")

    result["meeting_date"] = st.text_input("Meeting Date", value=result["meeting_date"])
    result["summary"] = st.text_area("Summary", value=result["summary"], height=80)

    st.markdown("**Decisions**")
    for i, d in enumerate(result["decisions"]):
        d["decision"] = st.text_input(f"Decision {i+1}", value=d["decision"], key=f"dec_{i}")
        d["context"] = st.text_input(f"Context {i+1}", value=d["context"], key=f"dec_ctx_{i}")

    st.markdown("**Action Items**")
    for i, a in enumerate(result["action_items"]):
        cols = st.columns([3, 2, 2, 1])
        a["task"] = cols[0].text_input(f"Task {i+1}", value=a["task"], key=f"task_{i}")
        a["owner"] = cols[1].text_input("Owner", value=a["owner"], key=f"owner_{i}")
        a["due_date"] = cols[2].text_input("Due", value=a["due_date"], key=f"due_{i}")
        a["priority"] = cols[3].selectbox(
            "Priority", ["high", "medium", "low"],
            index=["high", "medium", "low"].index(a["priority"]),
            key=f"pri_{i}"
        )

    st.markdown("**Open Risks**")
    for i, r in enumerate(result["open_risks"]):
        cols = st.columns([3, 2])
        r["risk"] = cols[0].text_input(f"Risk {i+1}", value=r["risk"], key=f"risk_{i}")
        r["mitigation_plan"] = cols[1].text_input(
            "Mitigation Plan", value=r["mitigation_plan"], key=f"mitigation_{i}"
        )

    st.divider()
    st.subheader("Send")

    col1, col2, col3 = st.columns(3)

    with col1:
        if st.button("Save to Excel tracker", help=f"Appends to {EXCEL_PATH} in this folder"):
            try:
                wb, dedup_report = build_workbook(result, path=EXCEL_PATH)
                wb.save(EXCEL_PATH)
                with open(EXCEL_PATH, "rb") as f:
                    st.session_state.tracker_bytes = f.read()
                st.session_state.dedup_report = dedup_report
                st.success(f"Saved to {EXCEL_PATH}")
            except Exception as e:
                st.error(f"Failed to save: {e}")

        if st.session_state.get("dedup_report"):
            for tab_name, key in [("Decisions", "decisions"), ("Action Items", "action_items"), ("Risks", "risks")]:
                r = st.session_state.dedup_report[key]
                if r["skipped"]:
                    st.caption(
                        f"{tab_name}: {r['added']} added, "
                        f"{len(r['skipped'])} skipped as already logged"
                    )
                    with st.expander(f"See what was skipped in {tab_name}"):
                        for item in r["skipped"]:
                            st.markdown(f"- {item}")
                else:
                    st.caption(f"{tab_name}: {r['added']} added")

        if st.session_state.get("tracker_bytes"):
            st.download_button(
                "Download Excel",
                data=st.session_state.tracker_bytes,
                file_name=EXCEL_PATH,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.caption("Save to the tracker first to enable download.")

    with col2:
        webhook_url = st.text_input("Slack webhook URL", type="password")
        if st.button("Post to Slack", disabled=not webhook_url):
            try:
                post_to_slack(webhook_url, result)
                st.success("Posted to Slack.")
            except Exception as e:
                st.error(f"Failed to post: {e}")

    with col3:
        st.download_button(
            "Download JSON",
            data=json.dumps(result, indent=2),
            file_name="extraction_output.json",
            mime="application/json"
        )
