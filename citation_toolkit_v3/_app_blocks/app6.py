    st.markdown("## Reference List Comparator")
    col1,col2=st.columns(2)
    with col1:
        st.markdown("**List A**")
        file_a=st.file_uploader("List A",type=["xml","docx","txt"],key="comp_a",label_visibility="collapsed")
    with col2:
        st.markdown("**List B**")
        file_b=st.file_uploader("List B",type=["xml","docx","txt"],key="comp_b",label_visibility="collapsed")
    st.markdown("**Manuscript (optional)**")
    ms_file=st.file_uploader("Manuscript",type=["docx"],key="comp_ms",label_visibility="collapsed")
    if st.button("Compare Lists",type="primary",disabled=not(file_a and file_b)):
        with st.spinner("Comparing..."):
            refs_a,label_a=load_ref_file(file_a); refs_b,label_b=load_ref_file(file_b)
            if refs_a and refs_b:
                all_c=[r["corpus"] for r in refs_a]+[r["corpus"] for r in refs_b]
                vec=TfidfVectorizer(ngram_range=(1,2),sublinear_tf=True,max_features=50000)
                vec.fit(all_c)
                emb_a=vec.transform([r["corpus"] for r in refs_a])
                emb_b=vec.transform([r["corpus"] for r in refs_b])
                matrix=cosine_similarity(emb_a,emb_b)
                matched,only_a,fuzzy,matched_b=[],[],[],set()
                for i,ra in enumerate(refs_a):
                    bj=int(matrix[i].argmax()); bs=float(matrix[i][bj])
                    if bs>=MATCH_THRESHOLD: matched.append((ra,refs_b[bj],bs)); matched_b.add(bj)
                    elif bs>=FUZZY_THRESHOLD: fuzzy.append((ra,refs_b[bj],bs))
                    else: only_a.append(ra)
                only_b=[refs_b[j] for j in range(len(refs_b)) if j not in matched_b]
                result=dict(matched=matched,only_in_a=only_a,only_in_b=only_b,fuzzy=fuzzy)
                usage={}
                if ms_file:
                    ms_doc=Document(io.BytesIO(ms_file.read()))
                    ms_paras=[p.text for p in ms_doc.paragraphs if p.text.strip()]
                    for ref in only_a+only_b+[x[0] for x in fuzzy]:
                        words=[w for w in ref["title"].split() if len(w)>5][:5]
                        found=[p[:200] for p in ms_paras if sum(1 for w in words if w.lower() in p.lower())>=min(3,len(words))]
                        usage[ref.get("id","")]=found
                st.session_state.update(dict(comp_result=result,comp_usage=usage,
                    comp_labels=(label_a,label_b),comp_refs=(refs_a,refs_b)))
    if st.session_state.comp_result:
        result=st.session_state.comp_result; usage=st.session_state.comp_usage
        label_a,label_b=st.session_state.comp_labels
        st.divider()
        col1,col2,col3,col4=st.columns(4)
        col1.metric("Matched",len(result["matched"])); col2.metric("Only in A",len(result["only_in_a"]))
        col3.metric("Only in B",len(result["only_in_b"])); col4.metric("Review",len(result["fuzzy"]))
        tabs=st.tabs([f"Only in A ({len(result['only_in_a'])})",f"Only in B ({len(result['only_in_b'])})",
                      f"Review ({len(result['fuzzy'])})",f"Matched ({len(result['matched'])})"])
        def ref_card(ref,color,locs=None):
            loc=""
            if locs: loc=f'<div style="font-size:0.75rem;color:#4fc3f7;margin-top:4px">Found in {len(locs)} paragraph(s)</div>'+"".join(f'<div style="font-size:0.74rem;color:#607080;font-style:italic">{l[:150]}...</div>' for l in locs[:2])
            st.markdown(f'<div class="ref-item {color}">{fmt_ref(ref)}{loc}</div>',unsafe_allow_html=True)
        with tabs[0]:
            if result["only_in_a"]:
                st.caption(f"In **{label_a}** but not **{label_b}**")
                for ref in result["only_in_a"]: ref_card(ref,"missing",usage.get(ref.get("id",""),[]))
            else: st.success("None")
        with tabs[1]:
            if result["only_in_b"]:
                st.caption(f"In **{label_b}** but not **{label_a}**")
                for ref in result["only_in_b"]: ref_card(ref,"warn",usage.get(ref.get("id",""),[]))
            else: st.success("None")
        with tabs[2]:
            for ra,rb,score in result["fuzzy"]:
                st.markdown(f'<div class="ref-item warning"><b>[{score:.3f}]</b> A:{fmt_ref(ra,True)}<br>B:{fmt_ref(rb,True)}</div>',unsafe_allow_html=True)
        with tabs[3]:
            for ra,rb,score in result["matched"]:
                st.markdown(f'<div class="ref-item ok">[{score:.3f}] {fmt_ref(ra,True)}</div>',unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# APP 4 UI — DOCUMENT MERGER
