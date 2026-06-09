"""
Placeholder to EndNote (page 11) - Citation Toolkit
Resolves citation placeholders against an .enlx library and rewrites them as
EndNote temporary citations {Author, Year #RecNum}, ready for Update Citations.

Three modes:
  * Typed / green-highlighted citations  -> document + library
  * [[REF #]] markers                    -> document + library + the numbered
        reference list those numbers belong to (pin it, since positions shift
        when the document is edited)
  * Both

Loaded via exec(open("page_placeholder_convert.py").read()); runs at module scope.
"""
import streamlit as st
import sys, os
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent) if "__file__" in globals() else os.getcwd())
import placeholder_convert as pc

if "pe_out" not in st.session_state:
    st.session_state["pe_out"] = None

M_TYPED = "Typed / green-highlighted citations"
M_REF = "[[REF #]] markers (from Bibliography Relink)"
M_BOTH = "Both"

HL = {"Green": "GREEN", "Bright green": "BRIGHT_GREEN", "Yellow": "YELLOW",
      "Turquoise": "TURQUOISE", "Pink": "PINK"}

st.markdown("## Placeholder \u2192 EndNote")
st.markdown(
    "Turn citation placeholders into EndNote temporary citations "
    "`{Author, Year #RecNum}` by resolving them against your library. Open the "
    "result in Word with the **same** library selected and run *Update Citations "
    "and Bibliography* (or feed it to **Citation Repair / Rebuild**)."
)

mode = st.radio("What are you converting?", [M_TYPED, M_REF, M_BOTH], key="pe_mode")
do_green = mode in (M_TYPED, M_BOTH)
do_ref = mode in (M_REF, M_BOTH)

st.divider()
c1, c2 = st.columns(2)
with c1:
    doc_up = st.file_uploader("Document with placeholders (.docx)", type=["docx"], key="pe_doc")
with c2:
    lib_up = st.file_uploader("EndNote library (.enlx)", type=["enlx"], key="pe_lib")

# ── [[REF #]] reference-list source ────────────────────────────────────────
bib_up = None
bib_from_doc = True
if do_ref:
    st.markdown("**Where do the `[[REF #]]` numbers come from?**")
    src = st.radio(
        "Reference-list source",
        ["This document's own numbered reference list",
         "A specific numbered reference list I upload (recommended if the document was edited)"],
        key="pe_bibsrc", label_visibility="collapsed")
    bib_from_doc = src.startswith("This document")
    if not bib_from_doc:
        bib_up = st.file_uploader("Numbered reference list (.docx)", type=["docx"], key="pe_bib")
    st.caption("`[[REF #]]` numbers are reference-list **positions** - they shift when "
               "references are added or removed. Pin the exact list those numbers were "
               "generated against (usually the draft you ran Bibliography Relink on).")

# ── options ────────────────────────────────────────────────────────────────
hl_label = "Green"
if do_green:
    hl_label = st.selectbox("Highlight color marking the typed citations", list(HL.keys()),
                            index=0, key="pe_hl")
apply_near = st.checkbox(
    "Also apply close matches where the year/journal differs (otherwise they are only suggested)",
    value=False, key="pe_near")

go = st.button("Convert", type="primary", key="pe_go")

if go:
    if not doc_up or not lib_up:
        st.warning("Upload both the document and the .enlx library.")
    elif do_ref and not bib_from_doc and not bib_up:
        st.warning("Upload the numbered reference list, or switch to using the document's own list.")
    else:
        try:
            with st.spinner("Matching against the library\u2026"):
                out, report = pc.convert(
                    doc_up.read(), lib_up.read(),
                    do_green=do_green, do_refmarkers=do_ref,
                    highlight=HL[hl_label], apply_near=apply_near,
                    bib_source_bytes=(bib_up.read() if (do_ref and not bib_from_doc and bib_up) else None),
                )
                report_docx = pc.build_report_docx(report)
                stem = os.path.splitext(doc_up.name)[0]
                st.session_state["pe_out"] = (out, report_docx, pc.summarize(report),
                                              report, f"{stem}_EndNote_ready.docx",
                                              f"{stem}_conversion_report.docx", do_ref)
        except Exception as e:
            st.error(f"Could not convert: {e}")

res = st.session_state.get("pe_out")
if res:
    out, report_docx, sm, report, out_name, rep_name, was_ref = res
    if was_ref and sm["placeholders"] and sm["resolved_refs"] == 0 and sm["unresolved_refs"]:
        st.warning("No `[[REF #]]` numbers resolved - the chosen reference list may not match "
                   "these markers (or has no numbered entries). Check the source list.")
    m = st.columns(4)
    m[0].metric("Placeholders", sm["placeholders"])
    m[1].metric("Applied refs", sm["resolved_refs"])
    m[2].metric("Suggested", sm["suggested_refs"])
    m[3].metric("Unresolved", sm["unresolved_refs"])

    d1, d2 = st.columns(2)
    with d1:
        st.download_button("\u2b07  Converted document (.docx)", data=out, file_name=out_name,
                           mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                           type="primary", use_container_width=True, key="pe_dl1")
    with d2:
        st.download_button("\u2b07  Match report (.docx)", data=report_docx, file_name=rep_name,
                           mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                           use_container_width=True, key="pe_dl2")

    if sm["suggested_refs"] or sm["unresolved_refs"]:
        st.info("Some references were only suggested or are missing from the library - "
                "see the match report. Suggested = surname matched but year/journal differs.")

    with st.expander("Match details", expanded=True):
        rows = []
        for r in report:
            applied = "; ".join("%s #%d" % (rec['sur'], rec['id']) for _, rec in r['resolved'])
            sugg = "; ".join("%s #%d?" % (rec['sur'], rec['id']) for _, rec in r['near'])
            miss = "; ".join(str(k) for k, rec in r['unresolved'] if rec is None)
            rows.append({"Placeholder": r['orig'][:50], "Applied": applied,
                         "Suggested": sugg, "Unresolved": miss})
        st.dataframe(rows, use_container_width=True, hide_index=True)
