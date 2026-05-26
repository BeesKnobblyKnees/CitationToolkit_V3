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
