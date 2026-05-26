"""
Citation Toolkit — main entry point.
Run with: streamlit run app.py
"""
import streamlit as st
from shared import APP_CSS

st.set_page_config(
    page_title="Citation Toolkit",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(APP_CSS, unsafe_allow_html=True)

# ── Sidebar navigation ────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown('''
    <div style="padding:1.4rem 0.5rem 0.6rem;">
      <div class="sidebar-logo"><em>Citation</em> Toolkit</div>
      <div class="sidebar-sub">Tachdjian\'s Pediatric Orthopaedics</div>
    </div>
    <div style="height:1px;background:var(--border);margin-bottom:1rem;"></div>
    ''', unsafe_allow_html=True)

    if "page" not in st.session_state:
        st.session_state["page"] = "citation_repair"

    PAGES = {
        "citation_repair":     ("01", "Citation Repair",        "Fix fields · Merge · Fill placeholders"),
        "library_tools":       ("02", "Library & Reference",    "PubMed · Compare lists · RecNums"),
        "bibliography_audit":  ("03", "Bibliography Auditor",   "Cross-ref vs published PDF"),
        "finalise":            ("04", "Finalise",               "Renumber · Health check · Rename"),
        "figure_inventory":    ("05", "Figure Inventory",       "Check figure names"),
        "figure_extractor":    ("06", "Figure Extractor",       "Crop & export figures from PDF/Word"),
    }

    def _nav(key):
        num, label, desc = PAGES[key]
        active = st.session_state["page"] == key
        if st.button(f"{num} · {label}", key=f"_nav_{key}", use_container_width=True):
            st.session_state["page"] = key
            st.rerun()

    st.markdown('<div class="sidebar-rule">Workflow</div>', unsafe_allow_html=True)
    for key in PAGES:
        _nav(key)

    st.markdown('<div style="height:1.5rem"></div>', unsafe_allow_html=True)

# ── Route to active page ──────────────────────────────────────────────────────
page = st.session_state["page"]

if   page == "citation_repair":    exec(open("page_citation_repair.py").read())
elif page == "library_tools":      exec(open("page_library_tools.py").read())
elif page == "bibliography_audit": exec(open("page_bibliography_audit.py").read())
elif page == "finalise":           exec(open("page_finalise.py").read())
elif page == "figure_inventory":   exec(open("page_figure_inventory.py").read())
elif page == "figure_extractor":   exec(open("page_figure_extractor.py").read())
