"""
Citation Repair — 02 · Broken Citation Fixer, 03 · Document Merger, 04 · Citation Repair as tabs.
"""
import streamlit as st, zipfile, re, base64, io, html
from pathlib import Path
from shared import *

st.markdown(APP_CSS, unsafe_allow_html=True)
st.markdown('<div class="app-label">A practical group &nbsp; 1 of 5</div>', unsafe_allow_html=True)
st.markdown("## Citation Repair")

st.markdown("""<style>
.stButton button { background:#8b1a1a !important; color:#faf7f2 !important; }
.stButton button p, .stButton button span, .stButton button div {
    color:#faf7f2 !important; font-weight:600 !important;
}
.stButton button:hover { background:#6b1212 !important; color:#faf7f2 !important; }
[data-testid="stExpander"] details summary {
    font-size:0 !important; overflow:hidden !important;
}
[data-testid="stExpander"] details summary p,
[data-testid="stExpander"] details summary span,
[data-testid="stExpander"] details summary div {
    font-size:0.92rem !important; color:#2e2416 !important;
    font-family:'Source Sans 3',sans-serif !important; font-weight:500 !important;
}
[data-testid="stExpander"] details summary svg {
    width:16px !important; height:16px !important; display:block !important;
}
</style>""", unsafe_allow_html=True)

st.markdown("Fix broken field codes, merge documents, and fill citation placeholders.")
st.divider()

# ── Initialize session state for all tabs ─────────────────────────────────────
for _k, _v in [
    ("fix_analysis",       None),
    ("fix_result",         None),
    ("fix_after_stage1",   None),
    ("fix_after_stage2",   None),
    ("fix_raw_xml",        None),
    ("fix_docx_bytes",     None),
    ("fix_doc_name",       ""),
    ("fix_karol_db_id",    None),
    ("fix_karol_rec_nums", None),
    ("fix_missing_refs",   None),
    ("merge_doc1",         None),
    ("merge_doc2",         None),
    ("merge_result",       None),
    ("repair_doc",         None),
    ("repair_lib",         None),
    ("repair_matches",     None),
    ("repair_review",      []),
    ("repair_done",        False),
    ("current_idx",        0),
    ("decisions",          []),
    ("doc_obj",            None),
    ("flagged",            []),
    ("mat",                None),
    ("refs",               []),
    ("vec",                None),
]:
    if _k not in st.session_state:
        st.session_state[_k] = _v

_tab1, _tab2, _tab3 = st.tabs(["02 · Broken Citation Fixer", "03 · Document Merger", "04 · Citation Repair"])

with _tab1:
        st.markdown('<div class="app-label">A practical tool &nbsp;02</div>', unsafe_allow_html=True)
        st.markdown("## Broken Citation Fixer")
        st.markdown('<div class="instruction-box">Use this when EndNote only recognizes some citations even though the bibliography shows all references — or when citations show as "Traveling Library" instead of your EndNote library.</div>', unsafe_allow_html=True)

        # Step 0 — Extract traveling library
        with st.expander("📥 Step 0 — Extract traveling library references (start here)", expanded=True):
            st.markdown('<div class="step-desc">If EndNote cannot find your refs, extract them from the Word file and import into your library.</div>', unsafe_allow_html=True)
            tl_file=st.file_uploader("Word document",type=["docx"],key="tl_doc")
            if tl_file:
                if st.button("Extract references",type="primary",key="tl_run"):
                    with st.spinner("Extracting..."):
                        try:
                            tl_xml,tl_count=extract_traveling_library_xml(tl_file.read())
                            st.success(f"Extracted {tl_count} references.")
                            st.download_button(f"⬇ Download EndNote XML ({tl_count} refs)",
                                data=tl_xml.encode('utf-8'),
                                file_name=Path(tl_file.name).stem+"_traveling_library.xml",
                                mime="application/xml",type="primary")
                            st.markdown('<div class="instruction-box">Import into your library: <b>File → Import → File</b> → select XML → Import Option: <code>EndNote XML</code> → Duplicates: <code>Discard Duplicates</code> → Import</div>',unsafe_allow_html=True)
                        except Exception as e: st.error(f"Error: {e}")

        st.divider()

        # Step 0b — Remap traveling library citations
        with st.expander("🔄 Step 0b — Remap citations still showing as Traveling Library"):
            st.markdown('<div class="step-desc">After importing the traveling library, some citations may still show as "Traveling Library" because their RecNums don\'t match your library\'s record IDs. This remaps them by author+year+title matching.</div>',unsafe_allow_html=True)
            col1,col2=st.columns(2)
            with col1: remap_doc=st.file_uploader("Word document",type=["docx"],key="remap_doc")
            with col2: remap_enl=st.file_uploader("EndNote library (.enl)",type=["enl"],key="remap_enl")
            if remap_doc and remap_enl:
                if st.button("Find and remap",type="primary",key="remap_run"):
                    with st.spinner("Comparing against your library..."):
                        try:
                            fixed_bytes,report=remap_traveling_citations(remap_doc.read(),remap_enl.read())
                            remapped=[r for r in report if r['status']=='remapped']
                            not_found=[r for r in report if r['status']=='not_found']
                            col1,col2,col3=st.columns(3)
                            col1.metric("Checked",len(report)); col2.metric("Remapped",len(remapped)); col3.metric("Not found",len(not_found))
                            if remapped:
                                st.success(f"Remapped {len(remapped)} citations.")
                                st.download_button("⬇ Download remapped document",data=fixed_bytes,
                                    file_name=Path(remap_doc.name).stem+"_remapped.docx",
                                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",type="primary")
                            if not_found:
                                st.warning(f"{len(not_found)} citations not found in your library — add them manually.")
                                with st.expander("Unmatched citations"):
                                    for r in not_found:
                                        st.markdown(f'<div class="ref-item error"><b>RecNum {r["old_rec_num"]}</b> — {r["author"]} ({r["year"]}) {r["title"]}</div>',unsafe_allow_html=True)
                        except Exception as e: st.error(f"Error: {e}")

        st.divider()

        # Stage 1
        stage1_done = st.session_state.fix_analysis is not None
        _c1 = ("done" if stage1_done else "active")
        _n1 = ("\u2713" if stage1_done else "STEP 1")
        _nc1 = ("done" if stage1_done else "")
        st.markdown(f'<div class="step-card {_c1}"><div class="step-header"><span class="step-num {_nc1}">{_n1}</span>'
            f' &nbsp; <span class="step-title">Repair broken citation field codes</span></div>'
            f'<div class="step-desc">Recovers citation XML data from the document&apos;s internal backup storage.</div></div>',
            unsafe_allow_html=True)
        if not stage1_done:
            doc_file=st.file_uploader("Word document (.docx)",type=["docx"],key="fix_doc_upload")
            if doc_file and st.button("Analyze & repair",type="primary"):
                with st.spinner("Scanning..."):
                    docx_bytes=doc_file.read()
                    analysis=analyze_docx_citations(docx_bytes)
                    fixed_xml,n_fixed=fix_broken_fields(analysis['raw'])
                    st.session_state.fix_analysis=analysis; st.session_state.fix_docx_bytes=docx_bytes
                    st.session_state.fix_after_stage1=fixed_xml; st.session_state.fix_raw_xml=fixed_xml
                    st.session_state.fix_doc_name=doc_file.name
                st.rerun()
        else:
            a=st.session_state.fix_analysis

            # Key metrics
            col1,col2,col3,col4 = st.columns(4)
            col1.metric("Total citation fields", a['total_fields'])
            col2.metric("EndNote currently sees",
                        a.get('endnote_sees', a['working']),
                        help="Unique references EndNote can read right now")
            col3.metric("Hidden in broken fields",
                        a.get('endnote_misses', a['broken_empty']),
                        delta=f"-{a.get('endnote_misses', a['broken_empty'])}" if a.get('endnote_misses', a['broken_empty']) > 0 else None,
                        delta_color="inverse",
                        help="References locked in broken field codes — EndNote cannot see these, causing bibliography undercount")
            col4.metric("Recovered by this fix", a['broken_empty'])

            if a.get('endnote_misses', 0) > 0:
                bib = a.get('bib_count', 0)
                sees = a.get('endnote_sees', a['working'])
                misses = a.get('endnote_misses', a['broken_empty'])
                st.error(
                    f"⚠ **EndNote is undercounting your references.** "
                    f"It can currently see **{sees}** unique references, but **{misses}** more "
                    f"are locked in broken citation field codes and invisible to EndNote. "
                    f"This is why your bibliography shows fewer entries than expected. "
                    f"Download the Stage 1 result below — this fix restores all {misses} hidden references."
                )
            elif a['broken_empty'] > 0:
                st.success(f"✓ {a['broken_empty']} broken field(s) recovered.")
            else:
                st.success("✓ All citation fields are intact — no broken fields found.")
            stage1_bytes=build_fixed_docx(st.session_state.fix_docx_bytes,st.session_state.fix_after_stage1)
            st.download_button("⬇ Download Stage 1 result",data=stage1_bytes,
                file_name=Path(st.session_state.fix_doc_name).stem+"_stage1.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

            # Orphaned superscript removal — run on the Stage 1 output
            st.markdown("---")
            st.markdown("**Optional: Remove orphaned superscripts**")
            st.markdown(
                "After a merge, plain-text superscript citation numbers sometimes remain "
                "next to working EndNote field codes (e.g. `135.40,47`). "
                "These are invisible to EndNote but visible to the reader. Click below to strip them."
            )
            if st.button("Remove orphaned plain superscripts"):
                with st.spinner("Scanning for orphaned superscripts..."):
                    cleaned_bytes, n_removed = remove_orphan_superscripts(stage1_bytes)
                if n_removed == 0:
                    st.info("No orphaned superscripts found — document is clean.")
                else:
                    st.success(f"✓ Removed {n_removed} orphaned superscript run(s).")
                    st.download_button(
                        "⬇ Download cleaned document",
                        data=cleaned_bytes,
                        file_name=Path(st.session_state.fix_doc_name).stem+"_cleaned.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        type="primary"
                    )

            if st.button("↺ Start over"):
                for k in ['fix_analysis','fix_raw_xml','fix_docx_bytes','fix_after_stage1','fix_after_stage2','fix_karol_db_id','fix_karol_rec_nums','fix_missing_refs']:
                    st.session_state[k]=defaults[k]
                st.rerun()

        st.divider()

        # Stage 2
        stage2_done=st.session_state.fix_karol_db_id is not None
        _s2_class = "done" if stage2_done else ("active" if stage1_done else "waiting")
        _s2_num   = "✓" if stage2_done else "STEP 2"
        _s2_nc    = "done" if stage2_done else ""
        st.markdown(f'''<div class="step-card {_s2_class}">
          <div class="step-header">
            <span class="step-num {_s2_nc}">{_s2_num}</span>
            <span class="step-title">Get your EndNote library fingerprint (db-id)</span>
          </div>
          <div class="step-desc">Export any refs from your library as XML (File → Export → XML), then upload below. The app extracts the library fingerprint automatically.</div>
        </div>''', unsafe_allow_html=True)
        if stage1_done and not stage2_done:
            col_a,col_b=st.columns(2)
            with col_a:
                xml_exp=st.file_uploader("EndNote XML export (gets db-id)",type=["xml"],key="fix_xml_export")
            with col_b:
                enl_f=st.file_uploader("EndNote library (.enl) — optional, checks for missing refs",type=["enl"],key="fix_enl")
            if xml_exp:
                db_id=extract_karol_db_id(xml_exp.read())
                if db_id:
                    st.session_state.fix_karol_db_id=db_id
                    st.success(f"✓ Library fingerprint found: `{db_id[:20]}...`"); st.rerun()
                else: st.error("No db-id found. Make sure you exported from EndNote as XML format.")
            if enl_f:
                with st.spinner("Reading library..."):
                    rns=get_karol_rec_nums(enl_f.read())
                    st.session_state.fix_karol_rec_nums=rns
                    missing=check_missing_from_karol(st.session_state.fix_raw_xml or "",set(rns.keys()))
                    st.session_state.fix_missing_refs=missing
                st.success(f"✓ {len(rns):,} refs in your library.")
                if missing: st.warning(f"⚠ {len(missing)} ref(s) not in your library: RecNums {', '.join(missing)}")
            with st.expander("Enter db-id manually"):
                mid=st.text_input("db-id",placeholder="e.g. s5pa559ekdxfr0esvw85...")
                if st.button("Use this db-id") and mid.strip():
                    st.session_state.fix_karol_db_id=mid.strip(); st.rerun()
        elif stage2_done:
            st.success(f"✓ Library db-id: `{st.session_state.fix_karol_db_id[:20]}...`")

        st.divider()

        # Stage 3
        stage3_done=st.session_state.fix_after_stage2 is not None
        stage3_ready=stage1_done and stage2_done
        _s3_class = "done" if stage3_done else ("active" if stage3_ready else "waiting")
        _s3_num   = "✓" if stage3_done else "STEP 3"
        _s3_nc    = "done" if stage3_done else ""
        st.markdown(f'''<div class="step-card {_s3_class}">
          <div class="step-header">
            <span class="step-num {_s3_nc}">{_s3_num}</span>
            <span class="step-title">Apply full fix and generate files</span>
          </div>
        </div>''', unsafe_allow_html=True)
        if stage3_ready and not stage3_done:
            missing=st.session_state.fix_missing_refs
            proceed=True
            if missing:
                st.warning(f"⚠ {len(missing)} ref(s) not in your library — add them manually first (EndNote → References → New Reference).")
                for rn in missing: st.markdown(f"- RecNum **{rn}**")
                proceed=st.checkbox("I've added the missing refs (or will add them later)")
            if proceed and st.button("Apply full fix",type="primary"):
                with st.spinner("Patching db-ids..."):
                    patched,_=patch_db_ids(st.session_state.fix_raw_xml,
                                           st.session_state.fix_analysis['db_ids'],
                                           st.session_state.fix_karol_db_id)
                    st.session_state.fix_after_stage2=patched; st.rerun()
        elif stage3_done:
            patched=st.session_state.fix_after_stage2
            final_bytes=build_fixed_docx(st.session_state.fix_docx_bytes,patched)
            working_after=len(re.findall(r'&lt;EndNote&gt;',patched))
            st.success("✓ Full fix applied.")
            col1,col2=st.columns(2)
            col1.metric("Citations now linked",working_after)
            col_d1,col_d2=st.columns(2)
            with col_d1:
                st.download_button("⬇ Download fixed document",data=final_bytes,
                    file_name=Path(st.session_state.fix_doc_name).stem+"_fully_fixed.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",type="primary")
            with col_d2:
                macro=generate_vba_macro(st.session_state.fix_doc_name)
                st.download_button("⬇ Download VBA macro",data=macro.encode(),
                    file_name="RelinkEndNoteCitations.bas",mime="text/plain")
            st.markdown('<div class="instruction-box"><b>Final step in Word:</b><br>1. Open the fixed document with your EndNote library connected<br>2. Open the .bas file in Notepad, copy all<br>3. In Word: Alt+F11 → Insert → Module → paste → Alt+F8 → RelinkAllCitations → Run<br>4. EndNote tab → Update Citations and Bibliography</div>',unsafe_allow_html=True)

    # ─────────────────────────────────────────────────────────────────────────────
    # APP 1 UI — CITATION REPAIR

with _tab2:
        st.markdown('<div class="app-label">A practical tool &nbsp;03</div>', unsafe_allow_html=True)
        st.markdown("## Document Merger")

        merge_mode = st.radio(
            "What do you need to do?",
            [
                "Restore lost citations — merge citation-intact OLD doc with text-edited NEW doc",
                "Repair already-merged document — fix broken field codes after a merge",
            ]
        )
        st.divider()

        if "Restore lost citations" in merge_mode:
            st.markdown('''<div class="instruction-box">
            <b>Use this when:</b> A merge broke or removed inline citations from the new document,
            but you still have the older version with all citations intact.<br><br>
            The app matches paragraphs between the two documents by text content and copies
            citation field codes from the old document into the new one wherever they are missing.
            All text edits in the new document are preserved.
            </div>''', unsafe_allow_html=True)

            col1, col2 = st.columns(2)
            with col1:
                st.markdown("**New document** — has text edits but missing/broken citations")
                new_file = st.file_uploader("New .docx", type=["docx"], key="merge_new")
            with col2:
                st.markdown("**Old document** — has all citations intact")
                old_file = st.file_uploader("Old .docx", type=["docx"], key="merge_old")

            if new_file and old_file:
                new_bytes_m = new_file.read()
                old_bytes_m = old_file.read()
                new_fname   = new_file.name

                new_analysis = analyze_merge_damage(new_bytes_m)
                col1, col2, col3 = st.columns(3)
                col1.metric("Bibliography entries",  new_analysis.get("bib_count", 0))
                col2.metric("Cited in text (new)",   new_analysis.get("cited_count", 0))
                col3.metric("Citations lost",         new_analysis.get("lost_in_merge", 0),
                            delta=f'-{new_analysis.get("lost_in_merge",0)}' if new_analysis.get("lost_in_merge",0) else None,
                            delta_color="inverse")

                if st.button("Restore citations from old document", type="primary"):
                    with st.spinner("Matching paragraphs and restoring citations..."):
                        try:
                            merged_bytes, rpt = safe_merge_documents(new_bytes_m, old_bytes_m)
                            st.success(
                                f"Done. Matched {rpt['matched']} paragraphs. "
                                f"Restored citations in {rpt['citations_restored']} paragraph(s). "
                                f"{rpt['unmatched']} paragraph(s) could not be matched to the old document."
                            )
                            col1,col2,col3,col4 = st.columns(4)
                            col1.metric("Matched",           rpt["matched"])
                            col2.metric("Citations restored",rpt["citations_restored"])
                            col3.metric("Already had cites", rpt["already_had_cites"])
                            col4.metric("Unmatched",         rpt["unmatched"])
                            if rpt["unmatched"] > 0:
                                st.info(
                                    f"{rpt['unmatched']} paragraph(s) are new content not in the old document "
                                    f"— these could not have citations restored automatically."
                                )
                            st.download_button(
                                "⬇ Download merged document",
                                data=merged_bytes,
                                file_name=Path(new_fname).stem + "_citations_restored.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                type="primary"
                            )
                            st.markdown('''<div class="instruction-box">
                            <b>After downloading:</b><br>
                            1. Open in Word with your EndNote library connected<br>
                            2. EndNote tab → Update Citations and Bibliography<br>
                            3. If any citations still unlinked → use App 2 Remap tool<br>
                            4. Check any unmatched paragraphs manually
                            </div>''', unsafe_allow_html=True)
                            if rpt["details"]:
                                with st.expander(f"Paragraphs with restored citations ({len(rpt['details'])})"):
                                    for d in rpt["details"]:
                                        st.markdown(
                                            f'<div class="ref-item ok">Para {d["para_idx"]+1} — {d["cites_added"]} citation run(s) restored<br><span style="font-size:0.8rem;font-style:italic">{d["text_preview"]}</span></div>',
                                            unsafe_allow_html=True
                                        )
                        except Exception as e:
                            st.error(f"Merge failed: {e}")
                            st.exception(e)

        else:
            st.markdown('''<div class="instruction-box">
            <b>When to use:</b> You merged two Word documents and EndNote no longer
            recognizes the citations. This tool accepts tracked changes safely
            (rescuing any citations inside deleted text) and repairs broken citation field codes.<br><br>
            <b>Best practice for future merges:</b> Before using Word's Compare, go to
            EndNote tab → Convert Citations → Convert to Unformatted Citations.
            This turns field codes into plain text like {Hall, 1997 #18} which survives
            merging perfectly. Then after accepting changes, use Update Citations and Bibliography.
            </div>''', unsafe_allow_html=True)

            col1, col2 = st.columns(2)
            with col1:
                st.markdown("**Post-merge document** — the merged file with broken citations")
                merged_file = st.file_uploader("Merged .docx", type=["docx"], key="merge_merged")
            with col2:
                st.markdown("**Original document** (optional) — used to detect lost citations")
                orig_file = st.file_uploader("Original .docx", type=["docx"], key="merge_orig")

            if merged_file:
                if st.button("Analyze & repair", type="primary"):
                    with st.spinner("Analyzing citation damage..."):
                        merged_bytes = merged_file.read()
                        orig_bytes   = orig_file.read() if orig_file else None
                        analysis     = analyze_merge_damage(merged_bytes)

                    st.markdown("### Damage report")
                    bib_count   = analysis.get("bib_count", 0)
                    cited_count = analysis.get("cited_count", 0)
                    lost_count  = analysis.get("lost_in_merge", 0)
                    field_count = analysis["total_en"]
                    sees        = analysis.get("endnote_sees", analysis["with_data"])
                    misses      = analysis.get("endnote_misses", analysis["empty_cite"])

                    col1,col2,col3,col4 = st.columns(4)
                    col1.metric("Bibliography entries",    bib_count,
                                help="Number of references in the numbered reference list")
                    col2.metric("EndNote currently sees",  sees,
                                help="Unique references EndNote can read from working field codes")
                    col3.metric("Hidden in broken fields", misses,
                                delta=f"-{misses}" if misses else None,
                                delta_color="inverse" if misses else "off",
                                help="References locked in broken field codes — EndNote cannot count these")
                    col4.metric("Lost during merge",       lost_count,
                                delta=f"-{lost_count}" if lost_count else None,
                                delta_color="inverse" if lost_count else "off",
                                help="References whose inline citations were removed by the merge")

                    if misses > 0:
                        st.error(
                            f"⚠ **EndNote is undercounting your references.** "
                            f"It currently sees **{sees}** unique references but **{misses}** more "
                            f"are locked in broken citation field codes and invisible to EndNote — "
                            f"causing your bibliography to show fewer entries than expected. "
                            f"This is fixed automatically below."
                        )
                    if lost_count > 0:
                        st.warning(
                            f"⚠ {lost_count} citation(s) were lost during the merge. "
                            f"Use **Restore lost citations** mode with the original document to recover them."
                        )
                    if misses == 0 and lost_count == 0 and bib_count > 0:
                        st.success(f"✓ All {bib_count} bibliography references are intact and visible to EndNote.")
                    if not analysis["balanced"]:
                        st.warning(
                            f"⚠ Unbalanced field markers "
                            f"(begin:{analysis['begins']}/sep:{analysis['separates']}/end:{analysis['ends']}) "
                            f"— some citation fields were split during merge."
                        )
                    elif field_count > 0:
                        st.success("✓ Citation field boundaries are intact.")
                    if len(analysis["db_ids"]) > 1:
                        st.info(f"Multiple library fingerprints found ({len(analysis['db_ids'])}) — citations from different libraries.")

                    with st.spinner("Repairing..."):
                        fixed_bytes, rpt = repair_post_merge_citations(merged_bytes, orig_bytes)
                    st.markdown("### Results")
                    col1,col2,col3 = st.columns(3)
                    col1.metric("Citations before", rpt["citations_before"])
                    col2.metric("Citations after",  rpt["citations_after"])
                    col3.metric("Steps applied",    len(rpt["steps"]))
                    for step in rpt["steps"]:
                        if step == "track_changes_accepted":
                            st.markdown("- ✓ Tracked changes accepted safely")
                        elif "restored" in step:
                            n = step.split("_")[1]
                            st.markdown(f"- ✓ {n} broken field(s) restored from backup data")
                        elif "lost" in step:
                            n = step.split("_")[0]
                            st.markdown(f"- ⚠ {n} citation(s) lost in merge — use Restore mode to recover")
                    if rpt.get("lost_rec_nums"):
                        st.warning(f"⚠ {len(rpt['lost_rec_nums'])} citation(s) from original not found after merge.")
                    st.download_button(
                        "⬇ Download repaired document",
                        data=fixed_bytes,
                        file_name=Path(merged_file.name).stem + "_repaired.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        type="primary"
                    )
                    st.markdown('''<div class="instruction-box">
                    <b>After downloading:</b><br>
                    1. Open in Word with your EndNote library connected<br>
                    2. EndNote tab → Update Citations and Bibliography<br>
                    3. Still issues? → Use App 2 Remap tool
                    </div>''', unsafe_allow_html=True)

with _tab3:
        st.markdown('<div class="app-label">A practical tool &nbsp;04</div>', unsafe_allow_html=True)
        st.markdown("## Citation Repair")
        st.markdown('<div class="step-desc">Find missing citation placeholders and match them to your EndNote library.</div>',unsafe_allow_html=True)
        col1,col2=st.columns(2)
        with col1: doc_file=st.file_uploader("Word document (.docx)",type=["docx"],key="doc_upload")
        with col2: lib_file=st.file_uploader("EndNote library (.xml)",type=["xml"],key="lib_upload")
        cfg1,cfg2=st.columns(2)
        with cfg1: mode=st.selectbox("Mode",["Interactive Review","Auto-Insert","Report Only"])
        with cfg2: use_pubmed=st.toggle("PubMed fallback",value=False)
        with st.expander("Custom citation markers"):
            custom=st.text_input("Additional markers (comma-separated)",placeholder="e.g. ??, [TBD]")
            if custom:
                extras=[re.escape(m.strip()) for m in custom.split(",") if m.strip()]
                CITATION_MARKERS=re.compile('|'.join(MISSING_PATTERNS+extras),re.IGNORECASE)
        run_btn=st.button("Scan Document",type="primary",disabled=not(doc_file and lib_file))
        if run_btn and doc_file and lib_file:
            for k in ["flagged","current_idx","decisions","doc_obj","refs","vec","mat","repair_done"]:
                st.session_state[k]=defaults[k]
            with st.spinner("Parsing..."):
                db=doc_file.read(); flagged,doc_obj=extract_flagged(db)
                refs=parse_endnote_xml_bytes(lib_file.read())
                if refs:
                    corpora=tuple(r["corpus"] for r in refs); vec,mat=build_tfidf(corpora)
                    st.session_state.update(dict(flagged=flagged,doc_obj=doc_obj,refs=refs,
                        vec=vec,mat=mat,current_idx=0,decisions=[],repair_done=False))
            if not refs: st.error("No references found in XML.")
            elif not flagged: st.warning("No citation markers found.")
            else: st.success(f"Found **{len(flagged)}** missing citation(s) across **{len(refs)}** refs.")
        flagged=st.session_state.flagged
        if flagged and not st.session_state.repair_done:
            st.divider()
            idx=st.session_state.current_idx; total=len(flagged); done=len(st.session_state.decisions)
            st.markdown(f'<div class="progress-outer"><div class="progress-inner" style="width:{int(done/total*100)}%"></div></div>',unsafe_allow_html=True)
            st.caption(f"{done} of {total} reviewed")
            if idx<total:
                item=flagged[idx]
                cands=match_sentence(item["sentence"],st.session_state.vec,st.session_state.mat,st.session_state.refs)
                best=cands[0]["score"] if cands else 0
                if mode=="Auto-Insert" and best>=TFIDF_THRESHOLD:
                    label=author_label(cands[0]["ref"]); para=st.session_state.doc_obj.paragraphs[item["para_idx"]]
                    insert_superscript(para,item["marker"],label)
                    st.session_state.decisions.append({**item,"action":"accepted","ref":cands[0]["ref"],"score":best,"candidates":cands})
                    st.session_state.current_idx+=1; st.rerun()
                elif mode=="Report Only":
                    st.session_state.decisions.append({**item,"action":"skipped","candidates":cands})
                    st.session_state.current_idx+=1; st.rerun()
                else:
                    st.markdown(f'<div class="match-card"><span class="match-marker">{item["marker"]}</span><div class="match-sentence">"{item["sentence"][:280]}"</div><div style="font-size:0.72rem;color:#3a4a5a">Para {item["para_idx"]+1}</div></div>',unsafe_allow_html=True)
                    st.markdown('<div class="section-label">Top Matches</div>',unsafe_allow_html=True)
                    chosen_idx=None
                    for j,c in enumerate(cands):
                        ca,cb=st.columns([1,8])
                        with ca:
                            if st.button("Use",key=f"pick_{idx}_{j}"): chosen_idx=j
                        with cb:
                            sc=score_class(c["score"])
                            st.markdown(f'<div style="display:flex;align-items:center;gap:8px;padding:5px 0"><span class="score-pill {sc}">{c["score"]:.3f}</span><span style="font-size:0.83rem;color:#c8d0db">{fmt_ref(c["ref"])}</span></div>',unsafe_allow_html=True)
                        if chosen_idx==j:
                            label=author_label(cands[chosen_idx]["ref"])
                            para=st.session_state.doc_obj.paragraphs[item["para_idx"]]
                            insert_superscript(para,item["marker"],label)
                            st.session_state.decisions.append({**item,"action":"accepted","ref":cands[chosen_idx]["ref"],"score":cands[chosen_idx]["score"],"candidates":cands})
                            st.session_state.current_idx+=1; st.rerun()
                    ba,bb=st.columns(2)
                    with ba:
                        if st.button("Skip",key=f"skip_{idx}"):
                            st.session_state.decisions.append({**item,"action":"skipped","candidates":cands})
                            st.session_state.current_idx+=1; st.rerun()
                    with bb:
                        if done>0 and st.button("Finish",key=f"fin_{idx}"):
                            st.session_state.repair_done=True; st.rerun()
            else:
                st.session_state.repair_done=True; st.rerun()
        if st.session_state.repair_done and st.session_state.decisions:
            st.divider()
            decisions=st.session_state.decisions
            accepted=sum(1 for d in decisions if d["action"]=="accepted")
            skipped=sum(1 for d in decisions if d["action"]=="skipped")
            col1,col2,col3=st.columns(3)
            col1.metric("Total",len(decisions)); col2.metric("Accepted",accepted); col3.metric("Skipped",skipped)
            cd1,cd2=st.columns(2)
            with cd1:
                st.download_button("⬇ Repaired document",data=doc_to_bytes(st.session_state.doc_obj),
                    file_name="manuscript_repaired.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",type="primary")
            with cd2:
                st.download_button("⬇ Decision report",data=write_repair_report(decisions),
                    file_name="citation_report.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

    # ─────────────────────────────────────────────────────────────────────────────
    # APP 3 UI — REFERENCE COMPARATOR
