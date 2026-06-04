"""
Citation Verifier Page (09)
Three-tier confidence check on a relinked document, keyed to the text BEFORE
each citation: (A) topical overlap with the reference title+keywords,
(B) whether the claim is retained published prose (vs the published PDF),
(C) optional PubMed abstract overlap. Triage, not certification.
Runs top-to-bottom via exec().
"""

import streamlit as st
import io
from pathlib import Path
from citation_verifier_module import verify

if "vf" not in st.session_state:
    st.session_state.vf = {"relinked": None, "pdf": None, "rows": None, "name": None}
vf = st.session_state.vf

st.title("Citation Verifier")
st.caption(
    "Checks each in-text citation against the text that PRECEDES it (how "
    "superscript citations work). Three tiers: (A) does the reference's topic "
    "match the claim, (B) is the claim retained published prose, (C) optional "
    "PubMed abstract check. This flags citations to review — it does not certify "
    "that a reference supports a claim; that still needs your judgment."
)

t1, t2 = st.tabs(["1 - Setup", "2 - Results"])

with t1:
    st.markdown("**Relinked document (.docx)** — output of Bibliography Relink (has field codes)")
    up = st.file_uploader("Relinked .docx", type=["docx"], key="vf_doc")
    if up:
        vf["relinked"] = up.read(); vf["name"] = up.name; st.success(f"Loaded {up.name}")

    st.markdown("**Published chapter PDF** — optional, enables tier B (context check)")
    up2 = st.file_uploader("Published PDF", type=["pdf"], key="vf_pdf")
    if up2:
        vf["pdf"] = up2.read(); st.success(f"Loaded {up2.name}")

    use_pm = st.checkbox("Enable tier C — PubMed abstract check (needs internet; slower)",
                         value=False, key="vf_pm")
    email = ""
    if use_pm:
        email = st.text_input("Your email (PubMed/NCBI requests this for API access)",
                              key="vf_email")
        st.caption("Adds ~1s per citation. NCBI etiquette: identify yourself with an email.")

    if vf["relinked"] and st.button("Run verification", type="primary", key="vf_run"):
        bar = st.progress(0.0, text="Checking citations…")
        def prog(i, n): bar.progress(min(1.0, (i + 1) / n), text=f"Citation {i+1} of {n}")
        rows = verify(vf["relinked"], vf["pdf"], use_pubmed=use_pm,
                      pubmed_email=email, progress=prog)
        vf["rows"] = rows
        bar.empty()
        st.success(f"Checked {len(rows)} citations. See the Results tab.")
    elif not vf["relinked"]:
        st.warning("Upload the relinked document to begin.")

with t2:
    if not vf["rows"]:
        st.warning("Run the verification in the Setup tab first.")
    else:
        rows = vf["rows"]
        from collections import Counter
        cnt = Counter(r["verdict"] for r in rows)
        a, b, c, d = st.columns(4)
        a.metric("Verified", cnt.get("verified", 0))
        b.metric("Supported", cnt.get("supported", 0))
        c.metric("Plausible", cnt.get("plausible", 0))
        d.metric("Review", cnt.get("review", 0))
        st.caption("Verified = retained published prose + on-topic reference. "
                   "Review = weak signals; read these yourself. Sort by verdict "
                   "and start with 'review'.")

        order = {"review": 0, "plausible": 1, "supported": 2, "verified": 3}
        table = [{
            "Verdict": r["verdict"],
            "Reference": r["reference"],
            "Claim (text before citation)": r["claim"],
            "A topical": r["topical_A"],
            "B published": r["published_B"],
            "C pubmed": r["pubmed_C"],
        } for r in sorted(rows, key=lambda r: order.get(r["verdict"], 9))]
        st.dataframe(table, use_container_width=True, hide_index=True)

        import csv, io as _io
        buf = _io.StringIO()
        w = csv.DictWriter(buf, fieldnames=list(table[0].keys()))
        w.writeheader(); w.writerows(table)
        st.download_button("Download verification table (.csv)",
                           data=buf.getvalue().encode("utf-8"),
                           file_name=Path(vf["name"] or "doc").stem + "_verification.csv",
                           mime="text/csv", key="vf_dl")
