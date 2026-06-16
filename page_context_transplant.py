"""
Context Transplant (page 12) - Citation Toolkit

Fix a section's broken citations by matching them, by surrounding WORDING, to a
linked twin of the same passage and transplanting its real EndNote field codes.
No bibliography or consistent numbering needed - the right tool when you're
working at the section level.

Loaded via exec(open("page_context_transplant.py").read()); runs at module scope.
"""
import streamlit as st
import sys, os
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent) if "__file__" in globals() else os.getcwd())
import context_transplant as ct

if "cx_out" not in st.session_state:
    st.session_state["cx_out"] = None

st.markdown("## Context Transplant")
st.markdown(
    "Repair a section whose citations are broken (EndNote **INVALID** errors, bare "
    "superscript numbers, or `[[REF n]]` markers) by matching each one - **by the "
    "wording around it** - to a linked twin of the same passage, and transplanting "
    "that twin's real field codes. No bibliography or numbering required. Open the "
    "result in Word and run *Update Citations and Bibliography* to renumber."
)

c1, c2 = st.columns(2)
with c1:
    tgt = st.file_uploader("Target - the working section (broken citations)", type=["docx"], key="cx_tgt")
with c2:
    srcf = st.file_uploader("Source - the linked twin (correct field codes)", type=["docx"], key="cx_src")

with st.expander("Options"):
    threshold = st.slider(
        "Match strictness (how much of the surrounding wording must agree)",
        min_value=0.2, max_value=0.8, value=0.4, step=0.05, key="cx_thr",
        help="Matching uses the distinctive words around each citation, so reworded sentences "
             "still match. 0.4 is a good default; lower it if the working text was edited heavily, "
             "raise it to demand closer wording.")
    fill_by_order = st.checkbox(
        "Fill any leftovers by citation order (best-effort)", value=False, key="cx_order",
        help="After wording-based matching, place any still-unmatched citations by their position "
             "between confident matches. Use only when the two versions have the same citations in "
             "the same order. These are flagged 'by position - verify' in the report.")
    replace_all = st.checkbox(
        "Also replace intact field codes (not just broken ones)", value=False, key="cx_all",
        help="Off by default - working citations that already have a valid record are left alone.")

go = st.button("Transplant", type="primary", key="cx_go")

if go:
    if not tgt or not srcf:
        st.warning("Upload both the target section and its linked twin.")
    else:
        try:
            with st.spinner("Matching citations by surrounding wording\u2026"):
                out, report = ct.transplant(tgt.read(), srcf.read(),
                                            threshold=threshold, replace_all=replace_all,
                                            fill_by_order=fill_by_order)
                rdoc = ct.build_report_docx(report)
                stem = os.path.splitext(tgt.name)[0]
                st.session_state["cx_out"] = (out, rdoc, ct.summarize(report), report,
                                              f"{stem}_transplanted.docx",
                                              f"{stem}_transplant_report.docx")
        except Exception as e:
            st.error(f"Could not transplant: {e}")

res = st.session_state.get("cx_out")
if res:
    out, rdoc, sm, report, out_name, rep_name = res
    if sm["citations"] == 0:
        st.info("No broken citations were found in the target. Intact field codes are left "
                "alone unless you tick \u201cAlso replace intact field codes\u201d.")
    m = st.columns(4)
    m[0].metric("Broken citations", sm["citations"])
    m[1].metric("By wording", sm["by_context"])
    m[2].metric("By position", sm["by_position"])
    m[3].metric("No match", sm["unmatched"])

    d1, d2 = st.columns(2)
    with d1:
        st.download_button("\u2b07  Transplanted document (.docx)", data=out, file_name=out_name,
                           mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                           type="primary", use_container_width=True, key="cx_dl1")
    with d2:
        st.download_button("\u2b07  Match report (.docx)", data=rdoc, file_name=rep_name,
                           mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                           use_container_width=True, key="cx_dl2")

    if sm["by_position"]:
        st.warning("%d citation(s) were placed **by position** rather than wording - check the "
                   "orange rows in the report and confirm they're right." % sm["by_position"])
    if sm["unmatched"]:
        st.info("%d citation(s) had no confident match (the wording differs too much between "
                "the two versions) and were left in place \u2013 see the red rows in the report. "
                "Typed author-year citations like these can go through Placeholder \u2192 EndNote."
                % sm["unmatched"])

    with st.expander("Match details", expanded=True):
        rows = [{"Context (run-up)": "\u2026" + r["context"][-48:], "Was": r["target"][:18],
                 "Now": ("%s #%s" % (r["matched"], r["recnum"]) if r["matched"] else "(left as-is)"),
                 "How": (r.get("method") or "-"), "Match": ("%.2f" % r["ratio"])} for r in report]
        st.dataframe(rows, use_container_width=True, hide_index=True)