"""
Chapter Page-Range Generator + Downloads (Citation Toolkit)

Upload a book PDF (with bookmarks) and for any chapter get:
  * total / text / bibliography page ranges (PDF pages AND printed book pages)
  * download the chapter as a standalone PDF
  * download the chapter as a Word document that maintains formatting, removes
    figures/tables/boxes/plates, and highlights every in-text figure/table/box/
    plate/video callout in BRIGHT CYAN (in-text reference superscripts preserved)

Loaded via exec(open("page_chapter_pages.py").read()); runs at module scope.
"""
import streamlit as st
import sys, os, tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent) if "__file__" in globals() else os.getcwd())
import chapter_pages as cp
import chapter_to_docx as ctd

st.markdown("## Chapter Page-Range Generator & Converter")
st.markdown(
    "Upload a book PDF and pick a chapter to get its **page ranges** (PDF and printed), "
    "download the **chapter PDF**, or convert it to a **Word document** — figures, tables, "
    "boxes, and plates removed, and every in-text figure/table/box/plate/video callout "
    "highlighted in bright cyan.")

pdf = st.file_uploader("Book PDF (with bookmarks)", type=["pdf"], key="cp_pdf")

@st.cache_data(show_spinner=False)
def _save(pdf_bytes):
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    tmp.write(pdf_bytes); tmp.close(); return tmp.name

@st.cache_data(show_spinner=False)
def _chapters(path): return cp.get_chapters(path)

@st.cache_data(show_spinner=False)
def _range(path, num): return cp.chapter_ranges(path, num)

@st.cache_data(show_spinner=False)
def _chapter_pdf(path, start, end): return ctd.extract_pdf(path, start, end)

@st.cache_data(show_spinner=False)
def _chapter_docx(path, num): return ctd.convert(path, num)

if pdf:
    path = _save(pdf.getvalue())
    with st.spinner("Reading bookmarks\u2026"):
        chapters = _chapters(path)
    if not chapters:
        st.error("No chapter bookmarks found in this PDF.")
    else:
        pick = st.selectbox("Chapter", [f"{c['num']} \u2014 {c['title']}" for c in chapters], key="cp_pick")
        num = int(pick.split(" \u2014 ")[0])

        with st.spinner("Reading chapter headers\u2026"):
            r = _range(path, num)

        if r:
            def fmt(sec):
                p = sec["pdf"]; pr = sec["printed"]
                return (f"{p[0]}\u2013{p[1]}" if p else "\u2014",
                        f"{pr[0]}\u2013{pr[1]}" if pr and pr[0] else "\u2014", sec["pages"])
            st.markdown(f"**Chapter {r['chapter']} \u2014 {r['title']}**")
            st.table([{"Part": n, "PDF pages": fmt(r[k])[0], "Printed pages": fmt(r[k])[1], "# pages": fmt(r[k])[2]}
                      for n, k in [("Total", "total"), ("Text", "text"), ("Bibliography", "bibliography")]])

            st.markdown("#### Downloads")
            c1, c2 = st.columns(2)

            with c1:
                if r["total"]["pdf"]:
                    s, e = r["total"]["pdf"]
                    if st.button("\U0001F4C4  Prepare chapter PDF", key="cp_pdf_btn"):
                        with st.spinner("Extracting chapter pages\u2026"):
                            st.session_state["cp_pdf_data"] = _chapter_pdf(path, s, e)
                    if st.session_state.get("cp_pdf_data"):
                        st.download_button("\u2b07  Download chapter PDF", st.session_state["cp_pdf_data"],
                                           file_name=f"Chapter_{num}_{r['title'].replace(' ', '_')}.pdf",
                                           mime="application/pdf", key="cp_pdf_dl")

            with c2:
                if st.button("\U0001F4DD  Convert to Word (figures/tables removed, callouts cyan)", key="cp_docx_btn"):
                    with st.spinner("Converting\u2026 (~20\u201340s for a long chapter)"):
                        try:
                            st.session_state["cp_docx_data"] = _chapter_docx(path, num)
                        except Exception as ex:
                            st.error(f"Conversion failed: {ex}")
                if st.session_state.get("cp_docx_data"):
                    st.download_button("\u2b07  Download Word document", st.session_state["cp_docx_data"],
                                       file_name=f"Chapter_{num}_{r['title'].replace(' ', '_')}.docx",
                                       mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                       key="cp_docx_dl")

            st.caption("Word conversion keeps headings, body text, Type I/II classification lists, and "
                       "in-text superscript references; removes figures, tables, boxes, and plates; and "
                       "highlights all in-text callouts in cyan. PDF page ranges are exact; a chapter's "
                       "opening printed page number can occasionally read \u00b11.")
