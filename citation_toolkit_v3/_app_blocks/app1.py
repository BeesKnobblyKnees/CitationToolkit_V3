    st.markdown("## PubMed Literature Search")
    st.markdown('<div class="instruction-box">Search PubMed for relevant articles. Results include abstracts and links to free full text where available (PubMed Central). Export results as EndNote XML to import directly into your EndNote library.</div>',unsafe_allow_html=True)
    query=st.text_input("Search query",placeholder="e.g. arthrogryposis clubfoot Ponseti treatment",help="Supports PubMed syntax: AND, OR, NOT, [MeSH], [ti], [au]")
    col1,col2,col3=st.columns(3)
    with col1: date_from=st.text_input("Year from",placeholder="2019")
    with col2: date_to=st.text_input("Year to",placeholder="2026")
    with col3: max_res=st.slider("Max results",5,50,20)
    journal_filter=st.text_input("Limit to journal (optional)",placeholder="e.g. J Pediatr Orthop")
    if st.button("Search PubMed",type="primary",disabled=not query.strip()):
        with st.spinner(f"Searching PubMed: {query}..."):
            results=pubmed_search_full(query.strip(),date_from.strip(),date_to.strip(),journal_filter.strip(),max_res)
        if not results: st.warning("No results found. Try broadening your search terms.")
        else:
            st.success(f"Found {len(results)} results.")
            xml_export=results_to_xml(results)
            st.download_button(f"⬇ Export all {len(results)} refs as EndNote XML",
                data=xml_export.encode('utf-8'),file_name="pubmed_results.xml",mime="application/xml")
            st.divider()
            for i,r in enumerate(results,1):
                authors_str='; '.join(r['authors'][:3])+(' et al.' if len(r['authors'])>3 else '')
                cit=f"{authors_str} ({r['year']}). *{r['journal']}*"
                if r['volume']: cit+=f" {r['volume']}"
                if r['issue']:  cit+=f"({r['issue']})"
                if r['pages']:  cit+=f":{r['pages']}"
                with st.expander(f"**{i}.** {r['title'][:100]}{'...' if len(r['title'])>100 else ''}"):
                    st.markdown(cit)
                    lc=st.columns(3)
                    with lc[0]: st.markdown(f"[PubMed]({r['pubmed_url']})")
                    with lc[1]:
                        if r['pmc_url']: st.markdown(f"[Free full text (PMC)]({r['pmc_url']})")
                        elif r['doi_url']: st.markdown(f"[DOI]({r['doi_url']})")
                        else: st.markdown("*No free full text*")
                    with lc[2]:
                        if r['doi']: st.caption(f"DOI: {r['doi']}")
                    if r['abstract']: st.markdown("**Abstract:**"); st.markdown(r['abstract'])
                    single=results_to_xml([r])
                    st.download_button("⬇ Export this ref",data=single.encode('utf-8'),
                        file_name=f"ref_{r['pmid']}.xml",mime="application/xml",key=f"exp_{r['pmid']}")

# ─────────────────────────────────────────────────────────────────────────────
# APP 8 UI — BATCH RENAME
