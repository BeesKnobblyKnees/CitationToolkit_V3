"""
Citation List (page 10) - Citation Toolkit
Generates an "in order of appearance" citation listing that mirrors EndNote 21's
Edit & Manage Citations dialog, as a downloadable Word document.

Loaded via exec(open("page_citation_listing.py").read()) from app.py, so this
file runs at module scope: no main(), no st.set_page_config().
"""
import streamlit as st
import sys, os
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent) if "__file__" in globals() else os.getcwd())
import citation_listing as cl

# ── session state ──────────────────────────────────────────────────────────
if "cl_result" not in st.session_state:
    st.session_state["cl_result"] = None      # (docx_bytes, totals, mode, fname)

st.markdown("## Citation List")
st.markdown(
    "Build a document listing every in-text citation **in order of appearance** - "
    "the same content as EndNote 21's *Edit & Manage Citations* view "
    "(citation group, references, count, library). Use it to walk the manuscript "
    "top to bottom and confirm each citation sits in the right place."
)

up = st.file_uploader(
    "Upload the chapter - a Word document (.docx) or the published PDF",
    type=["docx", "pdf"], key="cl_upload",
)

c1, c2 = st.columns([2, 1])
with c1:
    library_name = st.text_input(
        "Library name (shown in the Library column)",
        value="compressed chap 24 LLD", key="cl_lib",
        help="EndNote doesn't store the library name in the document, so set it here.",
    )
with c2:
    st.write("")
    st.write("")
    go = st.button("Generate listing", type="primary", use_container_width=True, key="cl_go")

st.caption(
    "Best results: a Word doc with **live EndNote fields** - that recovers record "
    "numbers, authors and years exactly (including large grouped citations). A PDF or "
    "a doc with bare superscript numbers can only give the citation **numbers** in order "
    "(no record numbers); if a numbered reference list is present it will be mapped in."
)

if go:
    if not up:
        st.warning("Upload a .docx or .pdf first.")
    else:
        data = up.read()
        name = up.name
        stem = os.path.splitext(name)[0]
        try:
            with st.spinner("Reading citations…"):
                if name.lower().endswith(".docx") and cl.has_endnote_fields(data):
                    res = cl.extract_from_docx_fieldcodes(data)
                    groups, totals = cl.normalize_fieldcode(res)
                    mode = "fieldcodes"
                    note = ""
                elif name.lower().endswith(".docx"):
                    res = cl.extract_superscripts_from_docx(data)
                    bib = cl.parse_numbered_bibliography(
                        cl.zipfile_read(data, "word/document.xml"))
                    groups, totals = cl.normalize_numbers(res, bib)
                    mode = "docx_numbers"
                    note = ("Degraded mode: this document has no live EndNote fields, so "
                            "only citation NUMBERS were recovered (no record numbers). "
                            "Verify against the source.")
                else:  # pdf
                    res = cl.extract_superscripts_from_pdf(data)
                    groups, totals = cl.normalize_numbers(res, res.get("bibliography"))
                    mode = "pdf_numbers"
                    note = ("Best-effort PDF mode: superscript citation numbers were "
                            "detected by font size/style and may be imperfect. No record "
                            "numbers are available from a PDF. Verify against the source.")

            if not groups:
                st.error("No in-text citations were found. If this is a PDF with unusual "
                         "superscript formatting, the linked Word document will give a far "
                         "better result.")
            else:
                docx_bytes = cl.build_listing_docx(
                    groups, totals, library_name=library_name,
                    source_name=name, mode_note=note,
                )
                st.session_state["cl_result"] = (docx_bytes, totals, mode,
                                                 f"{stem}_Citations_In_Order.docx", groups)
        except Exception as e:
            st.error(f"Could not process the file: {e}")

res = st.session_state.get("cl_result")
if res:
    docx_bytes, totals, mode, out_name, groups = res
    m1, m2, m3 = st.columns(3)
    m1.metric("Citation groups", totals["groups"])
    m2.metric("Citations", totals["citations"])
    m3.metric("References", totals["references"])

    st.download_button(
        "⬇  Download Word document", data=docx_bytes, file_name=out_name,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        type="primary", use_container_width=True, key="cl_dl",
    )

    with st.expander("Preview (first 25 groups)", expanded=True):
        rows = []
        for g in groups[:25]:
            rows.append({"#": g["order"], "Citation": g["display"], "Reference": "", "Count": ""})
            for r in g["rows"]:
                rows.append({"#": "", "Citation": "", "Reference": r["label"], "Count": r["count"]})
        st.dataframe(rows, use_container_width=True, hide_index=True)
