    st.markdown('<div class="app-label">A practical tool &nbsp;11</div>', unsafe_allow_html=True)
    st.markdown("## RecNum Inspector")
    st.markdown('''<div class="instruction-box">
    Reads a Word document that is connected to your EndNote library and shows you every
    reference's RecNum (Record Number), author, year, and title — so you can identify
    the correct RecNums to use when fixing field codes.<br><br>
    <b>Also:</b> if you have field codes using the wrong RecNums (e.g. placeholder RecNums
    500-506 we assigned before knowing your library's actual numbers), paste a remapping
    table and this tool will fix all field codes in one step and download the corrected document.
    </div>''', unsafe_allow_html=True)

    rn_file = st.file_uploader(
        "Word document (.docx) — must be library-connected (have working EN.CITE fields)",
        type=["docx"], key="rn_doc")

    if rn_file:
        raw_bytes = rn_file.read()
        with zipfile.ZipFile(io.BytesIO(raw_bytes)) as z:
            rn_xml = z.read('word/document.xml').decode('utf-8')

        # Extract all unique RecNum -> author/year/title from instrText
        import html as _html
        rn_data = {}
        for m in re.finditer(
                r'ADDIN EN\.CITE (&lt;EndNote&gt;[\s\S]+?)(?=</w:instrText>)', rn_xml):
            decoded = _html.unescape(m.group(1))
            rns     = re.findall(r'<RecNum>(\d+)</RecNum>', decoded)
            authors = re.findall(r'<Author>([^<]+)</Author>', decoded)
            years   = re.findall(r'<Year>([^<]+)</Year>', decoded)
            titles  = re.findall(r'<title>([^<]+)</title>', decoded)
            db_ids  = re.findall(r'db-id="([^"]*)"', decoded)
            for i, rn in enumerate(rns):
                if rn not in rn_data:
                    auth  = authors[i].split(',')[0] if i < len(authors) else (authors[0].split(',')[0] if authors else '?')
                    yr    = years[i] if i < len(years) else (years[0] if years else '?')
                    title = titles[0][:60] if titles else ''
                    db    = db_ids[i] if i < len(db_ids) else (db_ids[0] if db_ids else '')
                    rn_data[rn] = {'author': auth, 'year': yr, 'title': title, 'db_id': db}

        # Also check fldData for hidden refs
        import base64 as _b64
        for b64r in re.findall(r'<w:fldData[^>]*>([\s\S+?]+?)</w:fldData>', rn_xml):
            b64 = b64r.replace('\r','').replace('\n','').replace(' ','')
            pad = (4-len(b64)%4)%4
            try:
                dec = _b64.b64decode(b64+'='*pad).decode('utf-8',errors='replace').replace('\x00','')
                if '<EndNote>' not in dec: continue
                rns    = re.findall(r'<RecNum>(\d+)</RecNum>', dec)
                auths  = re.findall(r'<Author>([^<]+)</Author>', dec)
                yrs    = re.findall(r'<Year>([^<]+)</Year>', dec)
                titles = re.findall(r'<title>([^<]+)</title>', dec)
                db_ids = re.findall(r'db-id="([^"]*)"', dec)
                for i, rn in enumerate(rns):
                    if rn not in rn_data:
                        auth  = auths[i].split(',')[0] if i < len(auths) else (auths[0].split(',')[0] if auths else '?')
                        yr    = yrs[i] if i < len(yrs) else (yrs[0] if yrs else '?')
                        title = titles[0][:60] if titles else ''
                        db    = db_ids[i] if i < len(db_ids) else (db_ids[0] if db_ids else '')
                        rn_data[rn] = {'author': auth, 'year': yr, 'title': title,
                                       'db_id': db, 'hidden': True}
            except: pass

        correct_db = '9veea52shxtee2e2wsbpwvf89wz55atsf52s'
        st.metric("Unique RecNums found", len(rn_data))

        # ── Search / filter ───────────────────────────────────────────────────
        search = st.text_input("Search by author, year, or title", placeholder="e.g. Guirguis or 2017")

        rows = []
        for rn, info in sorted(rn_data.items(), key=lambda x: int(x[0]) if x[0].isdigit() else 9999):
            if search:
                needle = search.lower()
                haystack = f"{info['author']} {info['year']} {info['title']}".lower()
                if needle not in haystack:
                    continue
            db_ok = info['db_id'] == correct_db if info['db_id'] else False
            rows.append({
                'RecNum':  rn,
                'Author':  info['author'],
                'Year':    info['year'],
                'Title':   info['title'][:55],
                'DB-ID ✓': '✓' if db_ok else ('— empty' if not info['db_id'] else '✗ wrong'),
                'Hidden':  '⚠ fldData only' if info.get('hidden') else '',
            })

        if rows:
            import pandas as _pd
            df = _pd.DataFrame(rows)
            st.dataframe(df, use_container_width=True, height=400,
                         column_config={
                             'RecNum':  st.column_config.TextColumn('RecNum', width=80),
                             'Author':  st.column_config.TextColumn('Author', width=120),
                             'Year':    st.column_config.TextColumn('Year', width=60),
                             'Title':   st.column_config.TextColumn('Title'),
                             'DB-ID ✓': st.column_config.TextColumn('DB-ID', width=80),
                             'Hidden':  st.column_config.TextColumn('Status', width=110),
                         })
            # Download as CSV
            csv_data = _pd.DataFrame([{
                'RecNum': r['RecNum'], 'Author': r['Author'],
                'Year': r['Year'], 'Title': r['Title']} for r in rows])
            st.download_button("⬇ Download RecNum list as CSV",
                               data=csv_data.to_csv(index=False),
                               file_name="recnum_list.csv", mime="text/csv")
        else:
            st.info("No matching references found.")

        # ── RecNum remapping ──────────────────────────────────────────────────
        st.divider()
        st.markdown("### Fix wrong RecNums in another document")
        st.markdown(
            "If a document uses placeholder RecNums (e.g. 500, 501…) that don't exist "
            "in your library, enter the remapping below — then upload the document to fix "
            "and download a corrected version."
        )
        st.markdown("**Remapping format** — one per line: `old_recnum -> new_recnum`  "
                    "e.g. `500 -> 79`")

        remap_text = st.text_area("RecNum remapping", height=200,
                                   placeholder="500 -> 79\n501 -> 112\n502 -> 116\n503 -> 413")

        fix_file = st.file_uploader(
            "Document to fix (.docx)", type=["docx"], key="rn_fix_doc")

        if remap_text.strip() and fix_file and st.button("Apply remapping", type="primary"):
            # Parse remap table
            remap = {}
            errors = []
            for line in remap_text.strip().splitlines():
                line = line.strip()
                if not line or line.startswith('#'): continue
                parts = re.split(r'\s*->\s*', line)
                if len(parts) != 2:
                    errors.append(f"Bad line: '{line}'")
                    continue
                old_rn, new_rn = parts[0].strip(), parts[1].strip()
                if not old_rn.isdigit() or not new_rn.isdigit():
                    errors.append(f"Non-numeric RecNum in: '{line}'")
                    continue
                remap[old_rn] = new_rn

            if errors:
                for e in errors: st.error(e)
            elif not remap:
                st.warning("No valid remapping lines found.")
            else:
                fix_bytes = fix_file.read()
                with zipfile.ZipFile(io.BytesIO(fix_bytes)) as z:
                    fix_xml   = z.read('word/document.xml').decode('utf-8')
                    fix_files = {n: z.read(n) for n in z.namelist()}

                applied = []
                for old_rn, new_rn in remap.items():
                    before = fix_xml.count(f'&lt;RecNum&gt;{old_rn}&lt;/RecNum&gt;')
                    if before == 0:
                        st.warning(f"RecNum {old_rn} not found in document — skipped")
                        continue
                    fix_xml = fix_xml.replace(
                        f'&lt;RecNum&gt;{old_rn}&lt;/RecNum&gt;',
                        f'&lt;RecNum&gt;{new_rn}&lt;/RecNum&gt;')
                    fix_xml = fix_xml.replace(
                        f'&lt;rec-number&gt;{old_rn}&lt;/rec-number&gt;',
                        f'&lt;rec-number&gt;{new_rn}&lt;/rec-number&gt;')
                    fix_xml = re.sub(
                        rf'(&lt;key[^&]*&gt;){old_rn}(&lt;/key&gt;)',
                        rf'\g<1>{new_rn}\2', fix_xml)
                    after = fix_xml.count(f'&lt;RecNum&gt;{new_rn}&lt;/RecNum&gt;')
                    applied.append(f"{old_rn} → {new_rn} ({before} citation(s) updated)")

                fix_files['word/document.xml'] = fix_xml.encode('utf-8')
                out_buf = io.BytesIO()
                with zipfile.ZipFile(out_buf, 'w', zipfile.ZIP_DEFLATED) as zout:
                    for n, d in fix_files.items(): zout.writestr(n, d)
                out_buf.seek(0)
                fixed_bytes = out_buf.read()

                new_rns = set(re.findall(r'&lt;RecNum&gt;(\d+)&lt;/RecNum&gt;', fix_xml))
                st.success(f"✓ Applied {len(applied)} remapping(s). Document now has {len(new_rns)} unique RecNums.")
                for msg in applied:
                    st.markdown(f"  • {msg}")
                st.download_button(
                    "⬇ Download remapped document",
                    data=fixed_bytes,
                    file_name=Path(fix_file.name).stem + "_remapped.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    type="primary"
                )
