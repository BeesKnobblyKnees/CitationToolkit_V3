"""
Citation Rebuild Page  (07)
Part of Citation Toolkit. Runs top-to-bottom via exec(), like the other pages.
Detects broken {Author, Year #RecNum} citations and rebuilds CWYW field codes
using metadata from an EndNote .enlx library (or a prior working document).
"""

import streamlit as st
import tempfile
from pathlib import Path

from citation_rebuild_module import (
    CitationDetector,
    LibraryExtractor,
    CitationRebuilder,
)

# -- Session state -------------------------------------------------------------
if "rb" not in st.session_state:
    st.session_state.rb = {
        "use_library": True,
        "broken_path": None,
        "source_path": None,      # library OR prior doc
        "citations": None,
        "metadata": None,
        "output_path": None,
        "report": None,
    }
rb = st.session_state.rb

# A persistent temp dir that survives Streamlit reruns (NOT a context manager).
if "rb_tmpdir" not in st.session_state:
    st.session_state.rb_tmpdir = tempfile.mkdtemp(prefix="citation_rebuild_")
TMP = Path(st.session_state.rb_tmpdir)


def _save_upload(uploaded_file) -> Path:
    """Write an uploaded file to the persistent temp dir and return its path."""
    dest = TMP / uploaded_file.name
    dest.write_bytes(uploaded_file.getbuffer())
    return dest


# -- Header --------------------------------------------------------------------
st.title("Citation Rebuild")
st.caption(
    "Detect broken citations (shown as plain text like {Author, Year #RecNum}) "
    "and rebuild working EndNote field codes."
)

tab_setup, tab_detect, tab_map, tab_build, tab_download = st.tabs(
    ["1 - Setup", "2 - Detect", "3 - Map", "4 - Rebuild", "5 - Download"]
)

# -- 1 Setup -------------------------------------------------------------------
with tab_setup:
    st.subheader("Choose your inputs")

    rb["use_library"] = st.radio(
        "Source of reference metadata",
        options=[True, False],
        format_func=lambda v: "EndNote library (.enlx)" if v else "Prior working document (.docx)",
        horizontal=True,
        index=0 if rb["use_library"] else 1,
    )

    col1, col2 = st.columns(2)

    with col1:
        if rb["use_library"]:
            st.markdown("**EndNote library (.enlx)**")
            lib = st.file_uploader("Upload .enlx", type=["enlx"], key="rb_lib")
            if lib is not None:
                rb["source_path"] = _save_upload(lib)
                st.success(f"Loaded {lib.name}")
        else:
            st.markdown("**Prior working document (.docx)**")
            prior = st.file_uploader("Upload prior .docx", type=["docx"], key="rb_prior")
            if prior is not None:
                rb["source_path"] = _save_upload(prior)
                st.success(f"Loaded {prior.name}")

    with col2:
        st.markdown("**Broken document (.docx)**")
        broken = st.file_uploader("Upload broken .docx", type=["docx"], key="rb_broken")
        if broken is not None:
            rb["broken_path"] = _save_upload(broken)
            st.success(f"Loaded {broken.name}")

    if rb["broken_path"] and rb["source_path"]:
        st.info("Both files loaded. Go to 2 - Detect.")
    else:
        st.warning("Upload both files to continue.")

# -- 2 Detect ------------------------------------------------------------------
with tab_detect:
    st.subheader("Detect broken citations")
    if not rb["broken_path"]:
        st.warning("Upload a broken document in 1 - Setup first.")
    else:
        if st.button("Scan document", key="rb_scan"):
            try:
                rb["citations"] = CitationDetector.detect_broken_citations(rb["broken_path"])
            except Exception as e:
                st.error(f"Scan failed: {e}")

        if rb["citations"] is not None:
            cites = rb["citations"]
            st.success(f"Found {len(cites)} broken citations.")
            by_rec = {}
            for c in cites:
                by_rec.setdefault(c.recnum, []).append(c)
            st.dataframe(
                [
                    {"RecNum": r, "Author": v[0].author, "Year": v[0].year, "Count": len(v)}
                    for r, v in sorted(by_rec.items(), key=lambda kv: int(kv[0]))
                ],
                use_container_width=True,
            )

# -- 3 Map ---------------------------------------------------------------------
with tab_map:
    st.subheader("Map to reference metadata")
    if not rb["citations"]:
        st.warning("Run the scan in 2 - Detect first.")
    elif not rb["use_library"]:
        st.info("Prior-document metadata extraction isn't wired up yet - use the .enlx library option for now.")
    else:
        if st.button("Load metadata from library", key="rb_map"):
            try:
                targets = sorted({c.recnum for c in rb["citations"]})
                rb["metadata"] = LibraryExtractor.extract_from_enlx(rb["source_path"], targets)
            except Exception as e:
                st.error(f"Metadata load failed: {e}")

        if rb["metadata"]:
            st.success(f"Loaded metadata for {len(rb['metadata'])} references.")
            st.dataframe(
                [
                    {"RecNum": r, "Author": ref.author, "Year": ref.year,
                     "Title": (ref.title[:70] + "...") if len(ref.title) > 70 else ref.title}
                    for r, ref in sorted(rb["metadata"].items(), key=lambda kv: int(kv[0]))
                ],
                use_container_width=True,
            )

# -- 4 Rebuild -----------------------------------------------------------------
with tab_build:
    st.subheader("Rebuild field codes")
    if rb["use_library"] and not rb["metadata"]:
        st.warning("Load metadata in 3 - Map first.")
    elif not rb["broken_path"]:
        st.warning("Upload a broken document in 1 - Setup first.")
    else:
        if st.button("Run rebuild", key="rb_run"):
            try:
                out = TMP / (Path(rb["broken_path"]).stem + "_FIXED.docx")
                output_path, report = CitationRebuilder.rebuild_document(
                    broken_doc_path=rb["broken_path"],
                    library_path=rb["source_path"] if rb["use_library"] else None,
                    prior_doc_path=None if rb["use_library"] else rb["source_path"],
                    output_path=out,
                )
                rb["output_path"] = output_path
                rb["report"] = report
            except Exception as e:
                st.error(f"Rebuild failed: {e}")

        if rb["report"]:
            rep = rb["report"]
            c1, c2, c3 = st.columns(3)
            c1.metric("Found", rep["citations_found"])
            c2.metric("Rebuilt", rep["citations_rebuilt"])
            rate = round(100 * rep["citations_rebuilt"] / max(rep["citations_found"], 1))
            c3.metric("Success", f"{rate}%")
            if rep.get("recnums_processed"):
                st.dataframe(
                    [{"RecNum": r, "Author": a} for r, a in
                     sorted(rep["recnums_processed"].items(), key=lambda kv: int(kv[0]))],
                    use_container_width=True,
                )
            if rep.get("errors"):
                for err in rep["errors"]:
                    st.warning(err)

# -- 5 Download ----------------------------------------------------------------
with tab_download:
    st.subheader("Download fixed document")
    if not rb["output_path"]:
        st.warning("Run the rebuild in 4 - Rebuild first.")
    else:
        out = Path(rb["output_path"])
        if out.exists():
            st.download_button(
                "Download fixed .docx",
                data=out.read_bytes(),
                file_name=out.name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key="rb_dl",
            )
            st.success("Open it in Word, then run a field update (select all, F9) to refresh numbering.")
        else:
            st.error("Output file not found - re-run the rebuild.")
