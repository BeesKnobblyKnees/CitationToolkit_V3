    st.markdown('<div class="app-label">A practical tool &nbsp;05</div>', unsafe_allow_html=True)
    st.markdown("## Bibliography Auditor")
    st.markdown('''<div class="instruction-box">
    Cross-reference your Word document bibliography against the published PDF — for a specific chapter.
    Specify page ranges so the app only reads the relevant pages rather than the whole book.<br><br>
    <b>Finds:</b>
    <ul style="margin:0.4rem 0 0 1rem">
      <li>Refs in the published PDF missing from the Word doc</li>
      <li>Refs in the Word doc not in the published PDF (new additions)</li>
      <li>Published PDF refs not cited in the published body text</li>
      <li>Published PDF refs not cited in the Word doc body text</li>
      <li>Generates an importable EndNote file (.enw) for any missing refs</li>
    </ul>
    </div>''', unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**Word document (.docx)**")
        audit_doc = st.file_uploader("Word .docx", type=["docx"], key="audit_doc")
    with col2:
        st.markdown("**Published PDF (full book or chapter)**")
        audit_pdf = st.file_uploader("Published PDF", type=["pdf"], key="audit_pdf")

    if audit_pdf:
        import pdfplumber as _plumber
        with _plumber.open(audit_pdf) as _pp:
            _total_pages = len(_pp.pages)
        st.caption(f"PDF has {_total_pages} pages total")

        st.markdown("**Specify page ranges** — leave blank to use the whole PDF")
        pr1, pr2 = st.columns(2)
        with pr1:
            st.markdown("Body text pages (chapter content)")
            body_p1 = st.number_input("From page", min_value=1, max_value=_total_pages,
                                       value=1, key="body_p1")
            body_p2 = st.number_input("To page",   min_value=1, max_value=_total_pages,
                                       value=_total_pages, key="body_p2")
        with pr2:
            st.markdown("Bibliography pages")
            bib_p1 = st.number_input("From page", min_value=1, max_value=_total_pages,
                                      value=max(1, _total_pages-5), key="bib_p1")
            bib_p2 = st.number_input("To page",   min_value=1, max_value=_total_pages,
                                      value=_total_pages, key="bib_p2")

    if audit_doc and audit_pdf:
        if st.button("Run bibliography audit", type="primary"):
            import pdfplumber as _plumber, re as _re
            from docx import Document as _Doc
            from lxml import etree as _et

            with st.spinner("Extracting bibliography from PDF pages "
                            f"{bib_p1}–{bib_p2}..."):
                # ── Extract PDF bibliography (specified pages only) ──────────
                pdf_bib_text = ""
                with _plumber.open(audit_pdf) as _pdf:
                    for pg in _pdf.pages[bib_p1-1:bib_p2]:
                        t = pg.extract_text()
                        if t: pdf_bib_text += t + "\n"

                pdf_refs = {}
                # Two-column layout: refs appear as "N.\t..." anywhere on line
                for m in _re.finditer(
                    r'(?:^|\s{2,})(\d+[a-z]?)\.\s*\x07?([^\n]{10,})',
                    pdf_bib_text, _re.MULTILINE):
                    num  = m.group(1)
                    text = _re.sub(r'\s+', ' ', m.group(2)).strip()
                    if len(text) > 15 and num not in pdf_refs:
                        pdf_refs[num] = text

            with st.spinner("Extracting body text from PDF pages "
                            f"{body_p1}–{body_p2}..."):
                # ── Extract PDF body text (specified pages only) ─────────────
                pdf_body_text = ""
                with _plumber.open(audit_pdf) as _pdf:
                    for pg in _pdf.pages[body_p1-1:body_p2]:
                        t = pg.extract_text()
                        if t: pdf_body_text += t + "\n"

                # Find all citation numbers in PDF body text
                # Superscripts in published text appear as plain numbers after words
                pdf_body_cited = set()
                for m in _re.finditer(
                    r'(?<=[a-z\.\,\)])(\d+[a-z]?)(?=[,\.\s\n]|$)',
                    pdf_body_text, _re.MULTILINE | _re.IGNORECASE):
                    val = m.group(1)
                    if val.rstrip('a-z').isdigit():
                        pdf_body_cited.add(val)

            with st.spinner("Reading Word document..."):
                # ── Extract Word bibliography ────────────────────────────────
                wdoc = _Doc(audit_doc)
                ref_pat = _re.compile(r'^\s*(\d+[a-z]?)[\.)\s]\s+(.+)')
                word_bib = {}
                for p in wdoc.paragraphs:
                    m2 = ref_pat.match(p.text.strip())
                    if m2: word_bib[m2.group(1)] = p.text.strip()

                # ── Word body superscript citation numbers ───────────────────
                with zipfile.ZipFile(io.BytesIO(audit_doc.getvalue())) as _z:
                    doc_xml = _z.read('word/document.xml').decode('utf-8')
                W_ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
                root_elem = _et.fromstring(doc_xml.encode('utf-8'))
                word_body_cited = set()
                for r in root_elem.iter(f'{{{W_ns}}}r'):
                    rpr = r.find(f'{{{W_ns}}}rPr')
                    if rpr is None: continue
                    va = rpr.find(f'{{{W_ns}}}vertAlign')
                    if va is None or va.get(f'{{{W_ns}}}val') != 'superscript': continue
                    t_el = r.find(f'{{{W_ns}}}t')
                    txt  = (t_el.text or '') if t_el is not None else ''
                    for part in _re.split(r'[,;\s]+', txt):
                        p2 = part.strip().rstrip('.')
                        if p2.isdigit(): word_body_cited.add(p2)

                cited_rns = set(_re.findall(r'&lt;RecNum&gt;(\d+)&lt;/RecNum&gt;', doc_xml))
                total_field_rns = len(cited_rns)

            # ── Compare bibliographies ───────────────────────────────────────
            def ref_key(text):
                text = _re.sub(r'^\d+[a-z]?[\.)\s]\s*', '', text).strip()
                parts = _re.split(r',|;', text)
                auth  = parts[0].strip().split()[-1].lower() if parts[0].strip() else ''
                yr_m  = _re.search(r'\b(19|20)\d{2}\b', text)
                yr    = yr_m.group(0) if yr_m else ''
                return f"{auth} {yr}"

            word_keys = {ref_key(v): k for k, v in word_bib.items()}
            pdf_keys  = {ref_key(v): k for k, v in pdf_refs.items()}

            missing_from_word = [(num, pdf_refs[num]) for key, num in pdf_keys.items()
                                  if key not in word_keys]
            extra_in_word     = [(num, word_bib[num])  for key, num in word_keys.items()
                                  if key not in pdf_keys]

            pdf_nums  = set(pdf_refs.keys())
            not_in_pdf_body  = {n for n in pdf_nums if n not in pdf_body_cited}
            not_in_word_body = {n for n in pdf_nums if n not in word_body_cited}

            # ── Results ──────────────────────────────────────────────────────
            st.divider()
            st.markdown("### Section 1 — Bibliography comparison")
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Published PDF refs",     len(pdf_refs))
            c2.metric("Word doc refs",          len(word_bib))
            c3.metric("Missing from Word",      len(missing_from_word),
                      delta=f"-{len(missing_from_word)}" if missing_from_word else None,
                      delta_color="inverse" if missing_from_word else "off")
            c4.metric("EndNote field codes",    total_field_rns)

            if missing_from_word:
                st.error(f"⚠ **{len(missing_from_word)} ref(s) in the published PDF are missing from the Word document.**")
                with st.expander(f"Missing from Word ({len(missing_from_word)})", expanded=True):
                    for num, text in sorted(missing_from_word, key=lambda x: int(x[0]) if x[0].isdigit() else 999):
                        st.markdown(f'<div class="ref-item error">#{num} — {text[:100]}</div>',
                                    unsafe_allow_html=True)
                # Generate ENW
                enw_lines = []
                for num, text in missing_from_word:
                    auth_m = _re.match(
                        r'([A-Za-z\'\-]+)\s+(?:[A-Z]+\.?\s+)?(.+?)\.\s+(.+?)\.\s+(\d{4})', text)
                    if auth_m:
                        enw_lines += ['%0 Journal Article',
                                      f'%A {auth_m.group(1)}',
                                      f'%T {auth_m.group(2).strip()}',
                                      f'%J {auth_m.group(3).strip()}',
                                      f'%D {auth_m.group(4)}', '']
                    else:
                        enw_lines += ['%0 Journal Article', f'%T {text[:120]}', '']
                if enw_lines:
                    st.download_button("⬇ Download missing refs as EndNote import (.enw)",
                                       data='\n'.join(enw_lines),
                                       file_name="missing_refs.enw", mime="text/plain")

            if extra_in_word:
                st.info(f"ℹ {len(extra_in_word)} ref(s) in Word doc not in published PDF (new additions):")
                for num, text in extra_in_word[:15]:
                    st.markdown(f"  • #{num} — {text[:80]}")

            if not missing_from_word and not extra_in_word:
                st.success("✓ Word doc and published PDF bibliographies match perfectly.")

            st.divider()
            st.markdown("### Section 2 — Citations in published PDF body text")
            st.caption(f"Checked pages {body_p1}–{body_p2}")
            c5, c6 = st.columns(2)
            c5.metric("PDF refs cited in published body", len(pdf_nums) - len(not_in_pdf_body))
            c6.metric("PDF refs NOT cited in published body", len(not_in_pdf_body),
                      delta=str(len(not_in_pdf_body)) if not_in_pdf_body else None,
                      delta_color="inverse" if not_in_pdf_body else "off",
                      help="In the bibliography but no citation number found in the chapter body text")

            if not_in_pdf_body:
                with st.expander(f"Not cited in published body ({len(not_in_pdf_body)})", expanded=True):
                    st.caption("These appear in the published bibliography but no citation number "
                               "was detected in the published body text pages specified.")
                    for n in sorted(not_in_pdf_body, key=lambda x: int(x) if x.isdigit() else 999):
                        ref_text = pdf_refs.get(n, 'Unknown ref')
                        st.markdown(f'<div class="ref-item warning">#{n} — {ref_text[:100]}</div>',
                                    unsafe_allow_html=True)
            else:
                st.success("✓ All published PDF refs appear to be cited in the published body text.")

            st.divider()
            st.markdown("### Section 3 — Citations in Word document body text")
            c7, c8 = st.columns(2)
            c7.metric("PDF refs cited in Word body", len(pdf_nums) - len(not_in_word_body))
            c8.metric("PDF refs NOT cited in Word body", len(not_in_word_body),
                      delta=str(len(not_in_word_body)) if not_in_word_body else None,
                      delta_color="inverse" if not_in_word_body else "off",
                      help="In the published bibliography but no inline superscript found in the Word body text")

            if not_in_word_body:
                with st.expander(f"Not cited in Word body ({len(not_in_word_body)})", expanded=True):
                    st.caption("These appear in the published bibliography but their reference "
                               "number was not found as an inline superscript in the Word document. "
                               "May indicate citations lost during merge, or renumbering issues.")
                    for n in sorted(not_in_word_body, key=lambda x: int(x) if x.isdigit() else 999):
                        ref_text = pdf_refs.get(n, 'Unknown ref')
                        st.markdown(f'<div class="ref-item error">#{n} — {ref_text[:100]}</div>',
                                    unsafe_allow_html=True)
            else:
                st.success("✓ All published PDF refs appear as inline citations in the Word document.")
