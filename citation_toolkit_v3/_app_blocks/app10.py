    st.markdown('<div class="app-label">A practical tool &nbsp;10</div>', unsafe_allow_html=True)
    st.markdown("## Batch Rename")
    st.markdown('<div class="instruction-box">Apply bulk find-and-replace across your Word document using an Excel naming sheet. Use for final editing to update chapter names, author names, figure labels, section titles, etc.<br><br><b>Excel format:</b> Two columns — <code>Old Name</code> (find) and <code>New Name</code> (replace). One row per replacement.</div>',unsafe_allow_html=True)
    col1,col2=st.columns(2)
    with col1: ren_doc=st.file_uploader("Word document (.docx)",type=["docx"],key="batch_doc")
    with col2: ren_excel=st.file_uploader("Excel naming sheet (.xlsx)",type=["xlsx","xls"],key="batch_excel")
    col3,col4=st.columns(2)
    with col3: match_case=st.toggle("Match case",value=False)
    with col4: whole_word=st.toggle("Whole word only",value=False)
    pairs=[]
    if ren_excel:
        with st.spinner("Reading naming sheet..."):
            pairs=load_rename_pairs(ren_excel.read())
        if pairs:
            st.success(f"✓ {len(pairs)} rename pairs loaded.")
            with st.expander("Preview"):
                col1,col2=st.columns(2)
                col1.markdown("**Find**"); col2.markdown("**Replace**")
                for old,new in pairs[:20]: col1.markdown(old); col2.markdown(new)
                if len(pairs)>20: st.caption(f"...and {len(pairs)-20} more")
        else: st.error("Could not read pairs. Check Excel has 'Old Name' and 'New Name' columns.")
    if ren_doc and ren_excel and pairs:
        if st.button("Apply batch rename",type="primary"):
            with st.spinner("Applying..."):
                fixed_bytes,report=batch_rename(ren_doc.read(),pairs,match_case,whole_word)
            replaced=[r for r in report if r['status']=='replaced']
            not_found=[r for r in report if r['status']=='not_found']
            total_changes=sum(r['count'] for r in replaced)
            col1,col2,col3=st.columns(3)
            col1.metric("Total pairs",len(report)); col2.metric("Replaced",len(replaced)); col3.metric("Not found",len(not_found))
            st.success(f"✓ {total_changes} replacement(s) across {len(replaced)} pairs.")
            st.download_button("⬇ Download renamed document",data=fixed_bytes,
                file_name=Path(ren_doc.name).stem+"_renamed.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",type="primary")
            if replaced:
                with st.expander(f"Replacements made ({len(replaced)})"):
                    for r in replaced:
                        st.markdown(f'<div class="ref-item ok"><b>{r["old"]}</b> → {r["new"]} <span style="float:right;font-size:0.75rem">{r["count"]}x</span></div>',unsafe_allow_html=True)
            if not_found:
                with st.expander(f"Not found ({len(not_found)})"):
                    for r in not_found:
                        st.markdown(f'<div class="ref-item error"><b>{r["old"]}</b> — not found in document</div>',unsafe_allow_html=True)
