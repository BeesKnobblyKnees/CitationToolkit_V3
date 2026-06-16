"""
Bibliography Relink Page (09)
Relink a text-edited draft (bare superscript citations) against a citation-intact
old document, matching by reference identity (author/year) via the draft's own
bibliography. Transplants real field codes where possible, assembles them from
real records otherwise, and drops [[REF n]] placeholders for the rest.
Runs top-to-bottom via exec(), like the other pages.
"""

import streamlit as st
import io
from pathlib import Path

from citation_bibrelink_module import relink, parse_bibliography, index_old_fieldcodes, build_placeholders_docx

if "br" not in st.session_state:
    st.session_state.br = {"draft": None, "old": None, "fixed": None,
                           "report": None, "phs": None, "draft_name": None,
                           "bib_src": None}
br = st.session_state.br

st.title("Bibliography Relink")
st.caption(
    "Relink a text-edited draft whose citations are bare superscript numbers, "
    "using the draft's own numbered bibliography plus a citation-intact old "
    "document. Matching is by reference identity (author + year), so it is "
    "immune to renumbering. Real field codes are transplanted or assembled from "
    "real records; anything unresolved becomes a [[REF n]] placeholder."
)

t1, t2, t3 = st.tabs(["1 - Upload", "2 - Analyse & Relink", "3 - Download"])

with t1:
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**Edited draft (.docx)** — superscript numbers + bibliography")
        up_d = st.file_uploader("Edited draft", type=["docx"], key="br_draft")
        if up_d:
            br["draft"] = up_d.read(); br["draft_name"] = up_d.name
            st.success(f"Loaded {up_d.name}")
    with c2:
        st.markdown("**Old document (.docx)** — citation-intact EndNote field codes")
        up_o = st.file_uploader("Old document", type=["docx"], key="br_old")
        if up_o:
            br["old"] = up_o.read(); st.success(f"Loaded {up_o.name}")

    st.markdown("**Reference list source** — where the citation *numbers* are defined")
    src_mode = st.radio(
        "Reference list source",
        ["Whole chapter - the edited draft contains its own numbered bibliography",
         "A section - use a separate reference list I upload"],
        key="br_bibsrc", label_visibility="collapsed")
    if src_mode.startswith("A section"):
        up_b = st.file_uploader("Chapter's numbered reference list (.docx)", type=["docx"], key="br_bibfile")
        if up_b:
            br["bib_src"] = up_b.read(); st.success(f"Loaded reference list: {up_b.name}")
        st.caption("Point this at the chapter's MASTER numbered reference list. The draft's "
                   "superscript numbers are resolved against it, so a cut-out section that has no "
                   "bibliography of its own can still be relinked. Use the chapter's real numbering, "
                   "not a section's locally regenerated 1-N list.")
    else:
        br["bib_src"] = None

    if not (br["draft"] and br["old"]):
        st.warning("Upload both documents to continue.")
    elif src_mode.startswith("A section") and not br["bib_src"]:
        st.warning("Section mode: upload the chapter's reference list, or switch to whole-chapter mode.")

with t2:
    if not (br["draft"] and br["old"]):
        st.warning("Upload both documents in tab 1 first.")
    else:
        if st.button("Analyse", key="br_analyse"):
            bib, bibtext = parse_bibliography(br["bib_src"] or br["draft"])
            fbs, cbi = index_old_fieldcodes(br["old"])
            st.session_state.br["_preview"] = (len(bibtext), len(fbs), len(cbi))
        prev = st.session_state.br.get("_preview")
        if prev:
            st.write(f"Reference list entries: **{prev[0]}** · "
                     f"old-doc citation groups: **{prev[1]}** · "
                     f"unique reference records available: **{prev[2]}**")
            if prev[0] == 0:
                st.warning("No numbered reference list found. If this draft is a section, "
                           "switch to section mode in tab 1 and upload the chapter's list - "
                           "otherwise every citation becomes a [[REF #]] placeholder.")

        construct = st.checkbox(
            "Assemble field codes for partial matches (max coverage). "
            "Uncheck for transplant-only (safest, fewer citations).",
            value=True, key="br_construct")

        if st.button("Relink", type="primary", key="br_relink"):
            with st.spinner("Resolving citations and building field codes…"):
                fixed, report, phs = relink(br["draft"], br["old"], construct=construct,
                                            bib_source_bytes=br["bib_src"])
                br["fixed"], br["report"], br["phs"] = fixed, report, phs

        if br["report"]:
            r = br["report"]
            a, b, c, d = st.columns(4)
            a.metric("Citation locations", r["locations"])
            b.metric("Transplanted (exact)", r["exact"])
            c.metric("Assembled", r["constructed"])
            d.metric("Placeholders", r["placeholders"])
            st.caption(
                f"In-text cites **{r['unique_refs_cited']}** unique references. "
                f"The bibliography holds **{r['bibliography_entries']}** entries; "
                f"**{r['orphan_bib_entries']}** are no longer cited (removed sections) "
                f"and will drop out when EndNote rebuilds the bibliography."
            )
            if r.get("verify_count"):
                st.warning(
                    f"**{r['verify_count']}** citation(s) were same-author-same-year cases the "
                    "matcher couldn't separate by reference text. They were still inserted with the "
                    "best-guess record — double-check these against the list below:"
                )
                for nz, tx in r.get("verify", []):
                    st.caption("• " + ", ".join(str(n) for n in nz) + "  —  " + tx[:90])
            if r.get("footnote_locations"):
                st.caption(
                    f"Includes **{r['footnote_locations']}** citation group(s) in "
                    f"footnotes (the 'References …' lists): "
                    f"**{r.get('footnote_exact', 0)}** transplanted, "
                    f"**{r.get('footnote_constructed', 0)}** assembled, "
                    f"**{r.get('footnote_placeholders', 0)}** placeholder(s). "
                    "Transplanted footnote cites show as superscript until you "
                    "Update Citations; EndNote re-renders them to match your style."
                )
            st.info("Go to tab 3 to download. Open in Word → Update Citations and "
                    "Bibliography. The 'assembled' ones occasionally get rejected by "
                    "EndNote; any that show as INVALID can be re-inserted from your "
                    "library (their records are valid). Search the doc for '[[REF' "
                    "to find placeholders.")

with t3:
    if not br["fixed"]:
        st.warning("Run the relink in tab 2 first.")
    else:
        base = Path(br["draft_name"] or "document").stem
        st.download_button("Download relinked .docx", data=br["fixed"],
                           file_name=base + "_relinked.docx",
                           mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                           type="primary", key="br_dl")
        if br["phs"]:
            ph_docx = build_placeholders_docx(br["phs"])
            st.download_button("Download placeholder list (.docx)",
                               data=ph_docx,
                               file_name=base + "_placeholders.docx",
                               mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                               key="br_dl_ph")
            lines = ["CITATION PLACEHOLDERS TO INSERT", "=" * 50, ""]
            for nums, text in br["phs"]:
                lines.append(f"[[REF {','.join(str(n) for n in nums)}]]  ->  {text}")
            st.download_button("Download placeholder list (.txt)",
                               data="\n".join(lines).encode("utf-8"),
                               file_name=base + "_placeholders.txt",
                               mime="text/plain", key="br_dl_ph_txt")
        st.caption("After downloading: open in Word with EndNote, Update Citations "
                   "and Bibliography, then insert the placeholders listed above.")
