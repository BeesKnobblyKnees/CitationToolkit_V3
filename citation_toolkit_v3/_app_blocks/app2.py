    st.markdown('<div class="app-label">A practical tool &nbsp;02</div>', unsafe_allow_html=True)
    st.markdown("## Broken Citation Fixer")
    st.markdown('<div class="instruction-box">Use this when EndNote only recognizes some citations even though the bibliography shows all references — or when citations show as "Traveling Library" instead of your EndNote library.</div>', unsafe_allow_html=True)

    # Step 0 — Extract traveling library
    with st.expander("Step 0 — Extract traveling library references (start here)", expanded=True):
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
    with st.expander("Step 0b — Remap citations still showing as Traveling Library"):
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
