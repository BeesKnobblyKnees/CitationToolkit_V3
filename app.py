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
    "library_tools":       ("02", "Library & Reference",    "PubMed · Compare · RecNums · Citation Checker"),
    "bibliography_audit":  ("03", "Bibliography Auditor",   "Cross-ref vs published PDF"),
    "finalise":            ("04", "Finalise",               "Renumber · Health check · Rename"),
    "figure_inventory":    ("05", "Figure Inventory",       "Check figure names"),
    "figure_extractor":    ("06", "Figure Extractor",       "Crop & export figures from PDF/Word"),
    "citation_rebuild":    ("07", "Citation Rebuild",       "Detect · Map · Rebuild broken citations"),
    "citation_bibrelink":  ("08", "Bibliography Relink",    "Relink superscript cites via bibliography + old doc"),
    "citation_verifier":   ("09", "Citation Verifier",     "Check each citation against the text before it (A/B/PubMed)"),
    "citation_listing":   ("10", "Citation List", "In-order list matching EndNote Edit & Manage"),
    "placeholder_convert": ("11", "Placeholder → EndNote", "Resolve placeholders & typed citations to {Author, Year #RecNum} from your library"),
    "find_refs":           ("13", "Find Missing References", "Look up unresolved refs on CrossRef/PubMed, build importable .enw"),
    "context_transplant": ("12", "Context Transplant", "Fix a section's citations from a linked twin by matching surrounding wording"),
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
elif page == "citation_rebuild":   exec(open("page_citation_rebuild.py").read())
elif page == "citation_bibrelink":  exec(open("page_citation_bibrelink.py").read())
elif page == "citation_verifier":  exec(open("page_citation_verifier.py").read())
elif page == "citation_listing":  exec(open("page_citation_listing.py").read())
elif page == "placeholder_convert":  exec(open("page_placeholder_convert.py").read())
elif page == "find_refs":            exec(open("page_find_refs.py").read())
elif page == "context_transplant": exec(open("page_context_transplant.py").read())
# ── Final expander fix (must be last so it wins the cascade) ──────────────────
st.markdown("""
<style>
/* The arrow icon that leaks "_arr": hide the stray text, keep a real svg if present */
.epifhcv2 { font-size: 0 !important; color: transparent !important; }
.epifhcv2 svg { width: 16px !important; height: 16px !important; font-size: 1rem !important;
                color: var(--ink-dim) !important; fill: var(--ink-dim) !important; }
</style>
""", unsafe_allow_html=True)
st.markdown("""
<style>
div[data-testid="stExpander"] details > summary,
div[data-testid="stExpander"] details > summary > span,
div[data-testid="stExpander"] details > summary > span > div,
div[data-testid="stExpander"] details > summary > span > span {
  text-indent: 0 !important;
}
div[data-testid="stExpander"] details > summary { min-height: 0 !important; height: auto !important; }
div[data-testid="stExpander"] details > summary::before { display: none !important; content: none !important; }
div[data-testid="stExpander"] details > summary > span {
  flex-direction: row !important;
  flex-wrap: nowrap !important;
  width: 100% !important;
  min-width: 100% !important;
}
div[data-testid="stExpander"] details > summary > span > div { white-space: normal !important; flex: 1 1 auto !important; }
div[data-testid="stExpander"] details > summary > span > span { display: none !important; }
</style>
""", unsafe_allow_html=True)
