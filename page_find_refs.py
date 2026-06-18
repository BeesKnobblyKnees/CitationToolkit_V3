"""
Find Missing References -> EndNote  (Citation Toolkit)

When the Placeholder -> EndNote converter (page 11) leaves references UNRESOLVED
(red) because they aren't in the .enlx library, paste them here. This page looks
each one up on CrossRef and builds an importable .enw (or .ris) file. Import that
into the chapter library, then re-run the converter and the placeholders resolve.

It auto-loads the unresolved list from the last conversion if page 11 stored it
in st.session_state["pe_missing_refs"]; you can also paste references manually.

Loaded via exec(open("page_find_refs.py").read()); runs at module scope.
Follows toolkit conventions: no page-config call, no main(), no expander
widgets, and only non-widget session keys are initialized.
"""
import streamlit as st
import sys, os
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent) if "__file__" in globals() else os.getcwd())
import ref_finder as rf

# ---- non-widget session state only -----------------------------------------
if "rf_results" not in st.session_state:
    st.session_state["rf_results"] = None

st.markdown("## Find Missing References \u2192 EndNote")
st.markdown(
    "Paste the references that the **Placeholder \u2192 EndNote** converter couldn't "
    "resolve (the red / unresolved ones). This looks each up on CrossRef and builds "
    "an **importable `.enw`** file. Import it into the chapter library, then re-run "
    "the converter \u2014 the placeholders will resolve against the new records."
)

# Pre-fill from the last conversion's unresolved list, if page 11 stored it.
prefill = ""
_missing = st.session_state.get("pe_missing_refs")
if _missing:
    prefill = "\n".join(str(m) for m in _missing if str(m).strip())
    st.caption("Loaded %d unresolved reference(s) from the last conversion. "
               "Edit freely \u2014 fuller text (authors, title, journal, year) matches best."
               % len(_missing))

st.markdown("**References to look up** \u2014 one per line")
text = st.text_area(
    "refs", value=prefill, height=200, label_visibility="collapsed",
    placeholder="Barinaga G, Beason AM, Gardner MP. Novel surgical approach to segmental "
                "bone transport... J Am Acad Orthop Surg. 2018;26(22):e477-e482.\n"
                "Smolle MA, et al. Bone transport nails for reconstruction of lower limb "
                "diaphyseal defects in patients with bone sarcomas. Wien Klin Wochenschr. 2025.",
)

c1, c2 = st.columns([2, 1])
with c1:
    mailto = st.text_input(
        "Contact email (recommended)", value="",
        help="Used for CrossRef's faster 'polite pool' and as the NCBI/PubMed "
             "contact. Free, no key needed. Leave blank to skip.")
with c2:
    rows = st.number_input("Candidates per reference", min_value=1, max_value=10, value=3, step=1)

use_pubmed = st.checkbox(
    "Fall back to PubMed when CrossRef finds nothing", value=True,
    help="Good for medical references CrossRef misses. Uses NCBI E-utilities (free).")

insecure = st.checkbox(
    "Skip TLS certificate verification", value=False,
    help="The app already retries without verification when the network blocks TLS "
         "(common on hospital / SSL-inspection Wi-Fi). Tick this only to skip the "
         "verified attempt from the start - it just makes lookups a touch faster on "
         "those networks. Lookups are public CrossRef/PubMed metadata.")

go = st.button("Search", type="primary")

if go:
    queries = [ln.strip() for ln in text.splitlines() if ln.strip()]
    if not queries:
        st.warning("Add at least one reference (one per line).")
    else:
        src = "CrossRef + PubMed fallback" if use_pubmed else "CrossRef"
        with st.spinner("Searching %s for %d reference(s)..." % (src, len(queries))):
            try:
                st.session_state["rf_results"] = rf.find_references(
                    queries, rows=int(rows), mailto=(mailto.strip() or None),
                    use_pubmed=use_pubmed, insecure=insecure)
            except Exception as e:
                st.session_state["rf_results"] = None
                st.error("Search failed: %s" % e)

results = st.session_state.get("rf_results")

if results:
    st.divider()
    st.markdown("### Matches")
    st.caption("Pick the correct match for each reference, untick any you don't want, "
               "then download the import file below. Always eyeball the match \u2014 "
               "CrossRef returns the closest hit, not a guaranteed one.")

    chosen = []          # field dicts selected for export
    report_rows = []     # rows for the CSV report
    n_found = n_none = n_err = 0

    for i, r in enumerate(results):
        st.markdown("**%d.** %s" % (i + 1, (r["query"][:160] + ("..." if len(r["query"]) > 160 else ""))))
        if r["status"] == "error":
            n_err += 1
            err = r.get("error") or "unknown"
            st.error("Lookup error: %s" % err)
            if "CERTIFICATE" in err.upper() or "SSL" in err.upper():
                st.caption("That's a TLS error from your network's SSL inspection - tick "
                           "**Skip TLS certificate verification** above and search again.")
            report_rows.append({"Input": r["query"], "Status": "error",
                                "Source": r.get("source", ""), "Included": "no"})
            continue
        if r["status"] == "notfound" or not r["candidates"]:
            n_none += 1
            st.warning("No match. Try adding the journal name or title, or add it to EndNote by hand.")
            report_rows.append({"Input": r["query"], "Status": "notfound", "Included": "no"})
            continue

        n_found += 1
        cands = r["candidates"]
        src_label = {"crossref": "CrossRef", "pubmed": "PubMed"}.get(r.get("source"), r.get("source") or "")
        if src_label:
            st.caption("via %s" % src_label)

        def _label(idx, cands=cands):
            f = cands[idx]
            auth = f["authors"][0].split(",")[0] if f["authors"] else "?"
            if len(f["authors"]) > 1:
                auth += " et al."
            sc = (" \u00b7 score %.0f" % f["score"]) if f.get("score") else ""
            ttl = f["title"][:90] + ("..." if len(f["title"]) > 90 else "")
            return "%s (%s) %s \u2014 %s%s" % (auth, f["year"] or "n.d.", f["journal"] or "", ttl, sc)

        pick = st.radio(
            "Match", options=list(range(len(cands))), format_func=_label,
            key="rf_pick_%d" % i, label_visibility="collapsed")
        f = cands[pick]
        meta = " | ".join(x for x in [
            "Vol %s" % f["volume"] if f["volume"] else "",
            "No %s" % f["issue"] if f["issue"] else "",
            "pp %s" % f["pages"] if f["pages"] else "",
            "DOI %s" % f["doi"] if f["doi"] else "",
            "PMID %s" % f.get("pmid") if f.get("pmid") else "",
        ] if x)
        if meta:
            st.caption(meta)
        include = st.checkbox("Include in import file", value=True, key="rf_inc_%d" % i)
        if include:
            chosen.append(f)
        report_rows.append({
            "Input": r["query"], "Status": "found", "Source": f.get("source", ""),
            "Authors": "; ".join(f.get("authors", [])), "Title": f.get("title", ""),
            "Journal": f.get("journal", ""), "Year": f.get("year", ""),
            "Volume": f.get("volume", ""), "Issue": f.get("issue", ""),
            "Pages": f.get("pages", ""), "DOI": f.get("doi", ""),
            "PMID": f.get("pmid", ""), "Included": "yes" if include else "no",
        })
        st.markdown("")  # spacer

    st.divider()
    st.markdown("##### Summary")
    st.write("%d matched \u00b7 %d no match \u00b7 %d error \u00b7 **%d selected for import**"
             % (n_found, n_none, n_err, len(chosen)))

    if chosen:
        enw = rf.build_enw(chosen).encode("utf-8")
        ris = rf.build_ris(chosen).encode("utf-8")
        d1, d2 = st.columns(2)
        with d1:
            st.download_button("Download .enw  (EndNote Import)", data=enw,
                               file_name="found_references.enw",
                               mime="application/x-research-info-systems",
                               type="primary")
        with d2:
            st.download_button("Download .ris", data=ris,
                               file_name="found_references.ris",
                               mime="application/x-research-info-systems")
        st.caption("In EndNote: **File \u2192 Import \u2192 File**, Import Option "
                   "**EndNote Import** (for .enw) or **Reference Manager (RIS)** (for .ris), "
                   "with the chapter library open. Then re-run Placeholder \u2192 EndNote.")
    else:
        st.info("Nothing selected yet \u2014 pick at least one match above.")

    if report_rows:
        csv_bytes = rf.build_report_csv(report_rows).encode("utf-8-sig")
        st.download_button("Download lookup report (.csv)", data=csv_bytes,
                           file_name="found_references_report.csv", mime="text/csv")
        st.caption("The report logs every input line, the chosen match, its DOI/PMID, "
                   "source, and whether you included it \u2014 a paper trail for the audit.")
