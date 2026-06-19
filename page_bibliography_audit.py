"""
Standalone page.
"""
import streamlit as st, zipfile, re, base64, io, html
from pathlib import Path
from shared import *


def _page_text_columns(pg):
    """Read a PDF page in human reading order, column-aware.

    pdfplumber's extract_text() walks a two-column page straight across both
    columns, fusing the left-column line with the right-column line beside it
    (so reference 1 gets glued to the tail of reference 25, etc.). This detects
    a central gutter; if the page is two-column it reads the LEFT column fully
    top-to-bottom, then the RIGHT column, rebuilding lines from word positions.
    Full-width / single-column pages fall back to normal extraction."""
    try:
        words = pg.extract_words(use_text_flow=False)
    except Exception:
        words = None
    if not words:
        return pg.extract_text() or ""
    W = float(pg.width); cx = W / 2.0; band = 0.05 * W
    spanning = sum(1 for w in words if w['x0'] < cx - band and w['x1'] > cx + band)
    if spanning > max(4, 0.04 * len(words)):
        cols = [words]                                   # full-width text
    else:
        cols = [[w for w in words if (w['x0'] + w['x1']) / 2 < cx],
                [w for w in words if (w['x0'] + w['x1']) / 2 >= cx]]
    out = []
    for col in cols:
        col = sorted(col, key=lambda w: (w['top'], w['x0']))
        line, cur = [], None
        for w in col:
            if cur is None or abs(w['top'] - cur) <= 3.0:
                line.append(w['text']); cur = w['top'] if cur is None else cur
            else:
                out.append(' '.join(line)); line = [w['text']]; cur = w['top']
        if line:
            out.append(' '.join(line))
    return '\n'.join(out)


def _pdf_columnar_text(pdf, p1, p2):
    """Column-aware text for an inclusive 1-based page range."""
    parts = []
    for pg in pdf.pages[p1 - 1:p2]:
        parts.append(_page_text_columns(pg))
    return '\n'.join(parts).replace('\x07', '').replace('\u00ad', '')


def _parse_refs(text, hard_cap=4000):
    """Parse a numbered reference list into [(section, local_num, text), ...].

    Reference starts are detected as a number + optional '.'/')' at the start of
    a line, followed by an author letter. The sequence is then validated: a marker
    is accepted only if it is the next expected number, OR a '1' that restarts the
    count (a new section). This (a) skips stray numbers / page fragments, (b) keeps
    wrapped references whole, and (c) handles chapters that number each section
    separately, starting at 1 for each."""
    cands = [(int(m.group(1)), m.start(), m.end())
             for m in re.finditer(
                 r'(?m)^[ \t]*(\d{1,3})[a-z]?[.)]?\s+(?=[A-Za-z])', text)]
    accepted = []                       # (section, num, marker_start, text_start)
    sec, expected = 1, 1
    GAP = 8                              # tolerate up to a few deleted refs in a row
    for num, mstart, tend in cands:
        if num == expected:
            accepted.append((sec, num, mstart, tend)); expected = num + 1
        elif num == 1 and expected > 1:          # count restarts -> new section
            sec += 1
            accepted.append((sec, 1, mstart, tend)); expected = 2
        elif expected < num <= expected + GAP:   # small forward jump = deleted ref(s)
            accepted.append((sec, num, mstart, tend)); expected = num + 1
        # otherwise a stray / page-fragment / backward number -> skip
        if expected > hard_cap:
            break
    out = []
    for i, (s, num, mstart, tend) in enumerate(accepted):
        end = accepted[i + 1][2] if i + 1 < len(accepted) else len(text)
        body = re.sub(r'\s+', ' ', text[tend:end]).strip()
        out.append((s, num, body))
    return out


def _ref_labels(entries):
    """Turn [(section, num, text)] into {label: text}. Labels are plain numbers
    for a single-section list (backward compatible) and 'section.number' when the
    list restarts per section, so duplicate local numbers never collide."""
    multi = len({s for s, _, _ in entries}) > 1
    refs, local = {}, {}
    for sec, num, text in entries:
        if len(text) <= 15:
            continue
        label = ("%d.%d" % (sec, num)) if multi else str(num)
        refs[label] = text
        local[label] = str(num)
    return refs, local, multi

# Only initialize non-widget session state keys
# (audit_doc and audit_pdf are widget keys — Streamlit manages them)
if "audit_result" not in st.session_state:
    st.session_state["audit_result"] = None

st.markdown(APP_CSS, unsafe_allow_html=True)
st.divider()

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
            with _plumber.open(audit_pdf) as _pdf:
                pdf_bib_text = _pdf_columnar_text(_pdf, bib_p1, bib_p2)

            # Numbered reference list, parsed by validating the 1,2,3... sequence
            # and restarting per section (column-aware text + sequential split =
            # no fused/half references, and section-restart numbering supported).
            pdf_refs, pdf_local, pdf_multi = _ref_labels(_parse_refs(pdf_bib_text))

        with st.spinner("Extracting body text from PDF pages "
                        f"{body_p1}–{body_p2}..."):
            # ── Extract PDF body text (specified pages only) ─────────────
            with _plumber.open(audit_pdf) as _pdf:
                pdf_body_text = _pdf_columnar_text(_pdf, body_p1, body_p2)

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
            # Word reference lists use "N Author" or "N. Author"; parse with the
            # same restart-aware logic so per-section numbering doesn't collide.
            word_text = "\n".join(p.text for p in wdoc.paragraphs)
            word_bib, word_local, word_multi = _ref_labels(_parse_refs(word_text))

            # ── Word body superscript citation numbers ───────────────────
            with zipfile.ZipFile(io.BytesIO(audit_doc.getvalue())) as _z:
                doc_xml = _z.read('word/document.xml').decode('utf-8')
            W_ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
            root_elem = _et.fromstring(doc_xml.encode('utf-8'))
            # Superscript citations can be marked either by a direct vertAlign or
            # by a character style (citsup, sup, Superscript, FootnoteReference...).
            # Detecting only vertAlign misses citsup-styled numbers and flags every
            # reference as uncited. (root.iter is recursive, so runs inside tracked
            # changes <w:ins>/<w:moveTo> are included.)
            _sup_styles = {'citsup', 'sup', 'superscript', 'footnotereference',
                           'endnotebibliography', 'endnotereference'}
            word_body_cited = set()
            for r in root_elem.iter(f'{{{W_ns}}}r'):
                rpr = r.find(f'{{{W_ns}}}rPr')
                if rpr is None:
                    continue
                va = rpr.find(f'{{{W_ns}}}vertAlign')
                rs = rpr.find(f'{{{W_ns}}}rStyle')
                is_sup = (va is not None and va.get(f'{{{W_ns}}}val') == 'superscript') or \
                         (rs is not None and (rs.get(f'{{{W_ns}}}val') or '').lower()
                          in _sup_styles)
                if not is_sup:
                    continue
                t_el = r.find(f'{{{W_ns}}}t')
                txt  = (t_el.text or '') if t_el is not None else ''
                if not _re.fullmatch(r'[\d,;\s\u2013\-]+', txt or ''):
                    continue                       # digits + separators only
                for part in _re.split(r'[,;\s\u2013\-]+', txt):
                    p2 = part.strip().rstrip('.')
                    if p2.isdigit():
                        word_body_cited.add(p2)

            cited_rns = set(_re.findall(r'&lt;RecNum&gt;(\d+)&lt;/RecNum&gt;', doc_xml))
            total_field_rns = len(cited_rns)

        # ── Compare bibliographies ───────────────────────────────────────
        def ref_key(text):
            """(surname, year). The surname is the words before the first author's
            initials block, normalised so 'DeLuca' == 'De Luca', particles are kept
            ('van der Linden'), and a 'Jr'/'Sr' suffix is ignored. Far more reliable
            than taking the last token of the first comma-segment, which on a
            single-author entry grabs the year."""
            text = _re.sub(r'^\d+[a-z]?[\.)\s]\s*', '', text).strip()
            m = _re.match(r'^(.*?)\s+(?:[A-Z]\.?){1,3}(?=[\s,.;]|$)', text)
            sur = m.group(1) if m else (text.split()[0] if text.split() else '')
            sur = _re.sub(r"[\s'\-.]", '', sur.lower())
            yrs = _re.findall(r'\b(?:19|20)\d{2}\b', text)
            return sur, (yrs[-1] if yrs else '')

        word_keys = {ref_key(v): k for k, v in word_bib.items()}
        pdf_keys  = {ref_key(v): k for k, v in pdf_refs.items()}

        missing_from_word = [(num, pdf_refs[num]) for key, num in pdf_keys.items()
                              if key not in word_keys]
        extra_in_word     = [(num, word_bib[num])  for key, num in word_keys.items()
                              if key not in pdf_keys]

        # Reconcile: a reference that's "missing" on one side and "extra" on the
        # other with the SAME surname + a shared significant title word is really
        # the same paper with a metadata difference (usually the year). Pull those
        # out of the missing/extra lists into a clearer "differs between sources".
        def _title_words(text):
            t = _re.sub(r'^\d+[a-z]?[\.)\s]\s*', '', text)
            t = _re.sub(r'^.*?(?:[A-Z]\.?){1,3}[\s,.;]', '', t, count=1)   # drop authors
            return {w for w in _re.findall(r'[a-z]{5,}', t.lower())}

        differs_between = []
        _still_missing, _still_extra = [], list(extra_in_word)
        for pnum, ptext in missing_from_word:
            psur, pyr = ref_key(ptext); ptw = _title_words(ptext)
            hit = None
            for i, (wnum, wtext) in enumerate(_still_extra):
                wsur, wyr = ref_key(wtext)
                if wsur == psur and ptw & _title_words(wtext):
                    hit = i; break
            if hit is not None:
                wnum, wtext = _still_extra.pop(hit)
                differs_between.append((pnum, ptext, wnum, wtext))
            else:
                _still_missing.append((pnum, ptext))
        missing_from_word, extra_in_word = _still_missing, _still_extra

        pdf_nums  = set(pdf_refs.keys())
        not_in_pdf_body  = {lab for lab in pdf_refs
                            if pdf_local[lab] not in pdf_body_cited}
        not_in_word_body = {lab for lab in pdf_refs
                            if pdf_local[lab] not in word_body_cited}

        def _lab_sort(lab):
            try:
                return tuple(int(p) for p in str(lab).split('.'))
            except ValueError:
                return (999,)

        # ── Results ──────────────────────────────────────────────────────
        st.divider()
        st.markdown("### Section 1 — Bibliography comparison")
        if pdf_multi or word_multi:
            _ns = len({lab.split('.')[0] for lab in pdf_refs}) if pdf_multi else \
                  len({lab.split('.')[0] for lab in word_bib})
            st.caption(f"↳ Section-restart numbering detected ({_ns} sections). "
                       "Refs are labelled section.number (e.g. 2.5) so duplicate "
                       "numbers across sections don't collide. Body-citation checks "
                       "below match on the local number within a section, so a number "
                       "reused across sections is treated as cited if it appears anywhere.")
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
                for num, text in sorted(missing_from_word, key=lambda x: _lab_sort(x[0])):
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

        if differs_between:
            st.warning(f"⚠ {len(differs_between)} ref(s) present in BOTH but with a "
                       f"metadata difference (same author + title, usually the year):")
            for pnum, ptext, wnum, wtext in differs_between:
                st.markdown(f'<div class="ref-item warning">PDF #{pnum} — {ptext[:90]}<br>'
                            f'Word #{wnum} — {wtext[:90]}</div>', unsafe_allow_html=True)

        if not missing_from_word and not extra_in_word and not differs_between:
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
                for n in sorted(not_in_pdf_body, key=_lab_sort):
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
                for n in sorted(not_in_word_body, key=_lab_sort):
                    ref_text = pdf_refs.get(n, 'Unknown ref')
                    st.markdown(f'<div class="ref-item error">#{n} — {ref_text[:100]}</div>',
                                unsafe_allow_html=True)
        else:
            st.success("✓ All published PDF refs appear as inline citations in the Word document.")
