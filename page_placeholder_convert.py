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
M_REF = "Numbered markers  [[REF #]] · [#] · (#) · \u00b9\u00b2 superscript · ref lists"
M_BOTH = "Both"

HL = {"Green": "GREEN", "Bright green": "BRIGHT_GREEN", "Yellow": "YELLOW",
      "Turquoise": "TURQUOISE", "Pink": "PINK"}
STYLE = {"[[REF #]]": "refmark", "[#]": "bracket", "(#)": "paren", "Superscript \u00b9\u00b2": "superscript", "Reference number lists": "numlist"}

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
    lib_up = st.file_uploader("EndNote library — upload the .enlx, or YourLibrary.Data/sdb/sdb.eni", type=["enlx", "eni", "enl"], key="pe_lib")

# ── numbered-marker styles + reference-list source ─────────────────────────
bib_up = None
bib_from_doc = True
ref_styles = ('refmark',)
sectioned_mode = False
if do_ref:
    picks = st.multiselect(
        "Which numbered markers should I detect?",
        list(STYLE.keys()), default=["[[REF #]]"], key="pe_styles")
    ref_styles = tuple(STYLE[p] for p in picks)
    if any(s in ("[#]", "(#)") for s in picks):
        st.caption("`[#]` and `(#)` are only converted when the number is actually in the "
                   "reference list, so stray items like `(99)` or an enumeration are left alone. "
                   "`(#)` is the riskier one - enable it only if your citations really use parentheses.")
    if "Superscript \u00b9\u00b2" in picks:
        st.caption("Superscript detects Vancouver/JAMA-style superscript numbers (\u00b9 \u00b2) and "
                   "resolves them by reference number. Unit and variable powers like cm\u00b2 or x\u00b2 "
                   "are skipped automatically; the inserted citation is placed on the normal baseline "
                   "so EndNote re-applies your superscript style on Update.")
    if "Reference number lists" in picks:
        st.caption("Reference number lists handles plain comma/range lists like "
                   "\u201cReferences 7, 8, 13, 17\u201319, 33\u201d \u2014 the format used in figure/table "
                   "footnotes. Ranges are expanded (17\u201319 \u2192 17, 18, 19) and the whole list "
                   "becomes EndNote citations.")
    st.caption("All of the above are now detected in **footnotes, endnotes and table cells**, "
               "not just the document body.")
    st.markdown("**Where do the citation numbers come from?**")
    src = st.radio(
        "Reference-list source",
        ["This document's own numbered reference list",
         "A specific numbered reference list I upload (recommended if the document was edited)"],
        key="pe_bibsrc", label_visibility="collapsed")
    bib_from_doc = src.startswith("This document")
    if not bib_from_doc:
        bib_up = st.file_uploader("Numbered reference list (.docx)", type=["docx"], key="pe_bib")
    st.caption("The numbers are reference-list **positions** - pin the exact list those numbers "
               "were generated against (the numbered bibliography at the bottom of the chapter).")
    st.markdown("**Numbering style**")
    sectioned_mode = st.checkbox(
        "Sectioned reference list — numbering restarts in each section "
        "(published chapters: Pelvis 1…, Hip 1…, Femur 1…)",
        key="pe_sectioned")
    st.caption("Turn this on when the same number (e.g. 38) appears in several sections. "
               "Each in-text number is then resolved against the reference list of the "
               "**section it sits in**, matched by the body's section headings.")

# ── options ────────────────────────────────────────────────────────────────
typed_detect = "highlight"
hl_label = "Green"
if do_green:
    dmode = st.radio(
        "How are the typed citations marked?",
        ["Highlighted (recommended)",
         "By pattern \u2013 any \u201c(\u2026 Year)\u201d parenthesis, no highlight needed",
         "Both"],
        key="pe_detect")
    typed_detect = {"Highlighted (recommended)": "highlight",
                    "By pattern \u2013 any \u201c(\u2026 Year)\u201d parenthesis, no highlight needed": "pattern",
                    "Both": "both"}[dmode]
    if typed_detect in ("highlight", "both"):
        hl_label = st.selectbox("Highlight color marking the typed citations",
                                list(HL.keys()), index=0, key="pe_hl")
    if typed_detect in ("pattern", "both"):
        st.caption("Pattern mode scans for any parenthesis containing a 4-digit year, but only "
                   "converts one when a reference inside it resolves to your library - so plain "
                   "asides like \u201c(termed rebound deformity)\u201d are left untouched. Skim the "
                   "orange/red flags afterward to catch the occasional false positive or miss.")
apply_near = True  # near/ambiguous matches are always applied now, then flagged to verify
st.caption(
    "Every citation with a library match is inserted as `{Author, Year #RecNum}` - including the "
    "matcher's best guess for near and same-author/year cases. Those are listed in the match report "
    "as **verify** so you can double-check the handful that need it, rather than hand-entering them.")

go = st.button("Convert", type="primary", key="pe_go")

if go:
    if not doc_up or not lib_up:
        st.warning("Upload both the document and the .enlx library.")
    elif do_ref and not bib_from_doc and not bib_up:
        st.warning("Upload the numbered reference list, or switch to using the document's own list.")
    elif do_ref and not ref_styles:
        st.warning("Pick at least one numbered-marker style to detect ([[REF #]], [#], or (#)).")
    else:
        try:
            with st.spinner("Matching against the library\u2026"):
                out, report = pc.convert(
                    doc_up.read(), lib_up.read(),
                    do_green=do_green, do_refmarkers=do_ref,
                    highlight=HL[hl_label], apply_near=apply_near,
                    typed_detect=typed_detect, ref_styles=ref_styles, sectioned=sectioned_mode,
                    bib_source_bytes=(bib_up.read() if (do_ref and not bib_from_doc and bib_up) else None),
                )
                report_docx = pc.build_report_docx(report)
                stem = os.path.splitext(doc_up.name)[0]
                missing_files = pc.build_missing_imports(report)
                st.session_state["pe_out"] = (out, report_docx, pc.summarize(report),
                                              report, f"{stem}_EndNote_ready.docx",
                                              f"{stem}_conversion_report.docx", do_ref,
                                              missing_files, stem)
        except Exception as e:
            st.error(f"Could not convert: {e}")

res = st.session_state.get("pe_out")
if res:
    out, report_docx, sm, report, out_name, rep_name, was_ref, missing_files, stem = res
    if was_ref and sm["placeholders"] and sm["resolved_refs"] == 0 and sm["unresolved_refs"]:
        st.warning("No numbered markers resolved - the chosen reference list may not match "
                   "these markers (or has no numbered entries). Check the source list.")
    m = st.columns(5)
    m[0].metric("Placeholders", sm["placeholders"])
    m[1].metric("Applied refs", sm["applied_refs"])
    m[2].metric("To verify", sm["verify_refs"])
    m[3].metric("Unresolved", sm["unresolved_refs"])
    m[4].metric("Dangling", sm.get("dangling_refs", 0))

    d1, d2 = st.columns(2)
    with d1:
        st.download_button("\u2b07  Converted document (.docx)", data=out, file_name=out_name,
                           mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                           type="primary", use_container_width=True, key="pe_dl1")
    with d2:
        st.download_button("\u2b07  Match report (.docx)", data=report_docx, file_name=rep_name,
                           mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                           use_container_width=True, key="pe_dl2")

    if missing_files:
        st.markdown("**Missing references** (not in your library) \u2014 import into EndNote, or upload "
                    "the .docx to the Find-Missing-References app (one clean reference per line, so it "
                    "won't split them into fragments).")
        e1, e2, e3 = st.columns(3)
        with e1:
            st.download_button("\u2b07  Missing refs (.ris)", data=missing_files["ris"],
                               file_name=f"{stem}_missing_refs.ris",
                               mime="application/x-research-info-systems",
                               use_container_width=True, key="pe_ris")
        with e2:
            st.download_button("\u2b07  Missing refs (.enw)", data=missing_files["enw"],
                               file_name=f"{stem}_missing_refs.enw",
                               mime="application/x-endnote-library",
                               use_container_width=True, key="pe_enw")
        with e3:
            st.download_button("\u2b07  For Ref-Finder (.docx)", data=missing_files["docx"],
                               file_name=f"{stem}_missing_refs.docx",
                               mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                               use_container_width=True, key="pe_mdocx")

    if sm["verify_refs"] or sm["unresolved_refs"] or sm.get("dangling_refs"):
        msg = ("All matched references were applied. **%d** are best-guess matches flagged to "
               "verify (the report marks same-author/year ones first); **%d** had no library match "
               "and are left red - add them to the library."
               % (sm["verify_refs"], sm["unresolved_refs"]))
        if sm.get("dangling_refs"):
            msg += (" **%d** dangling - the in-text number has no entry in the reference list at all "
                    "(reference deleted or numbering off); fix the list or remove the citation."
                    % sm["dangling_refs"])
        st.info(msg)

    rows = []
    for r in report:
        applied = "; ".join("%s #%d" % (rec['sur'], rec['id'])
                            for _, rec in (r['resolved'] + r['suggested']))
        amb = set(r.get('ambiguous', []))
        verify = "; ".join(("%s #%d (same yr)" % (rec['sur'], rec['id'])) if key in amb
                           else ("%s #%d" % (rec['sur'], rec['id']))
                           for key, rec in r['suggested'])
        miss = "; ".join(str(k) for k, _ in r['missing'])
        dang = "; ".join("no ref #%s" % n for n in r.get('dangling', []))
        rows.append({"Placeholder": r['orig'][:50], "Applied": applied,
                     "Verify": verify, "Unresolved": miss, "Dangling": dang})
    st.dataframe(rows, use_container_width=True, hide_index=True)
