"""
placeholder_convert.py  --  turn citation placeholders into EndNote temporary
citations ({Author, Year #RecNum}) by resolving them against an .enlx library.

Placeholder kinds:
  * Typed author-year citations, detected either by HIGHLIGHT (a colored marker
    tells the app "this is a citation") or by PATTERN (any "(... 4-digit-year ...)"
    parenthesis). Pattern mode needs no highlighting but only touches a parenthesis
    when at least one reference inside it resolves to the library, so ordinary
    asides like "(termed rebound deformity)" or "(see 2019 update)" are left alone.
  * [[REF n, n, n]] markers (no highlight needed; found by text pattern), each
    number resolved via a numbered reference list (number -> author/year).

Output states (per reference):
  * applied    -> {Author, Year #RecNum}, no highlight
  * suggested  -> surname matched but year/journal differs; original text kept on
                  an ORANGE background, not auto-applied unless apply_near=True
  * unresolved -> not in the library; original text kept on a RED background
"""
import io, re, zipfile, sqlite3, difflib, html, copy
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

try:
    from citation_bibrelink_module import parse_bibliography
except Exception:
    parse_bibliography = None

ORANGE = "FFC000"   # suggested
RED = "FF5B5B"      # unresolved
_CITE_RE = re.compile(r'\(([^()]*\b(?:18|19|20)\d{2}\b[^()]*)\)')


# ── library ────────────────────────────────────────────────────────────────
def load_library(enlx_bytes):
    zf = zipfile.ZipFile(io.BytesIO(enlx_bytes))
    eni = next((n for n in zf.namelist() if n.endswith('sdb.eni')), None)
    if not eni:
        raise ValueError("No sdb.eni database found inside the .enlx file.")
    raw = zf.read(eni)
    import tempfile
    tmp = tempfile.NamedTemporaryFile(suffix='.eni', delete=False)
    tmp.write(raw); tmp.close()
    con = sqlite3.connect(tmp.name); con.row_factory = sqlite3.Row
    def clean(s): return re.sub(r'<[^>]+>', '', s or '').replace('\r', ' / ').replace('\n', ' ').strip()
    lib = []
    for r in con.execute('SELECT id,trash_state,author,year,title,secondary_title FROM refs'):
        if (r['trash_state'] or 0) != 0:
            continue
        au = clean(r['author']); first = au.split(' / ')[0] if au else ''
        sur = first.split(',')[0].strip()
        lib.append({'id': r['id'], 'sur': sur, 'surl': sur.lower(), 'year': str(r['year'] or ''),
                    'jour': clean(r['secondary_title']), 'title': clean(r['title'])})
    con.close()
    return lib


def match_to_library(surname, year, journal_hint, lib):
    sl = (surname or '').lower()
    jtoks = [t.lower() for t in re.findall(r'[A-Z][a-z]{2,}', journal_hint or '')]
    def jbonus(r):
        return 0.3 if jtoks and any(t in r['jour'].lower() for t in jtoks) else 0.0
    best = None
    for r in lib:
        if not year or r['year'] != year:
            continue
        ratio = difflib.SequenceMatcher(None, sl, r['surl']).ratio()
        starts = bool(sl) and (r['surl'].startswith(sl[:4]) or sl.startswith(r['surl'][:4]))
        score = ratio + (0.25 if starts else 0) + jbonus(r)
        if score >= 0.85 and (best is None or score > best[0]):
            best = (score, r)
    if best:
        return best[1], best[0], 'exact'
    near = None
    for r in lib:
        ratio = difflib.SequenceMatcher(None, sl, r['surl']).ratio()
        starts = bool(sl) and (r['surl'].startswith(sl[:5]) or sl.startswith(r['surl'][:5]))
        if ratio >= 0.8 or starts:
            score = ratio + (0.2 if starts else 0) + jbonus(r) + 0.2
            if near is None or score > near[0]:
                near = (score, r)
    if near and near[0] >= 0.9:
        return near[1], near[0], 'near'
    return None, 0.0, 'none'


def _token(rec):
    return "%s, %s #%d" % (rec['sur'], rec['year'], rec['id'])


def _parse_ref_text(text):
    s = text.strip().lstrip('.').lstrip('(').strip()
    ym = re.search(r'\b(?:18|19|20)\d{2}\b', s)
    year = ym.group(0) if ym else ''
    sm = re.match(r"([A-Za-z'\u00C0-\u017F\-]+)", s)
    sur = sm.group(1) if sm else ''
    jh = s.split(year)[-1] if year else s
    return sur, year, jh


def _is_green(run, color='GREEN'):
    try:
        return run.font.highlight_color is not None and color in str(run.font.highlight_color)
    except Exception:
        return False


def _has_shd(run):
    rpr = run._r.find(qn('w:rPr'))
    return rpr is not None and rpr.find(qn('w:shd')) is not None


def _shade_run(run, fill):
    rpr = run._r.get_or_add_rPr()
    for old in rpr.findall(qn('w:shd')):
        rpr.remove(old)
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear'); shd.set(qn('w:color'), 'auto'); shd.set(qn('w:fill'), fill)
    rpr.append(shd)


def _new_like(paragraph, ref_run, text, fill=None):
    """New run that inherits ref_run's font (minus highlight/shading), optional fill."""
    nr = paragraph.add_run(text)
    nr.font.highlight_color = None
    src = ref_run._r.find(qn('w:rPr'))
    if src is not None:
        ex = nr._r.find(qn('w:rPr'))
        if ex is not None:
            nr._r.remove(ex)
        clone = copy.deepcopy(src)
        for tag in ('w:highlight', 'w:shd'):
            for e in clone.findall(qn(tag)):
                clone.remove(e)
        nr._r.insert(0, clone)
    if fill:
        _shade_run(nr, fill)
    return nr


def _insert_like(paragraph, after_r, ref_run, text, fill=None):
    nr = _new_like(paragraph, ref_run, text, fill)
    after_r.addnext(nr._r)
    return nr._r


# ── numbered citation markers: [[REF n]] / [n] / (n) ──────────────────────── #
_NUM_PAT = {
    'refmark': re.compile(r'\[\[REF\s*([\d,\s\u2013-]+)\]\]'),
    'bracket': re.compile(r'\[\s*(\d{1,3}(?:\s*[,\u2013-]\s*\d{1,3})*)\s*\]'),
    'paren':   re.compile(r'\(\s*(\d{1,3}(?:\s*[,\u2013-]\s*\d{1,3})*)\s*\)'),
}
_FMT = {'refmark': '[[REF %s]]', 'bracket': '[%s]', 'paren': '(%s)'}


def _expand_nums(inner):
    out = []
    for part in re.split(r'\s*,\s*', inner.strip()):
        rng = re.match(r'(\d+)\s*[\u2013-]\s*(\d+)$', part)
        if rng:
            out += list(range(int(rng.group(1)), int(rng.group(2)) + 1))
        elif part.isdigit():
            out.append(int(part))
    return out


def _fmt_marker(style, nums):
    return _FMT[style] % ', '.join(str(n) for n in nums)


def _splice_marker(para, s, e, segments):
    """Replace the [s,e) span of a paragraph's text (which may cross runs) with
    the given (text, fill) segments, preserving surrounding runs/formatting."""
    runs = para.runs
    texts = [r.text or '' for r in runs]
    pos, spans = 0, []
    for t in texts:
        spans.append((pos, pos + len(t))); pos += len(t)
    ri_s = next((ri for ri, (a, b) in enumerate(spans) if a <= s < b), None)
    ri_e = next((ri for ri, (a, b) in enumerate(spans) if a < e <= b), None)
    if ri_s is None or ri_e is None:
        return
    prefix = texts[ri_s][:s - spans[ri_s][0]]
    suffix = texts[ri_e][e - spans[ri_e][0]:]
    ref = runs[ri_s]
    runs[ri_s].text = prefix
    runs[ri_s].font.highlight_color = None
    for ri in range(ri_s + 1, ri_e + 1):
        runs[ri].text = ''
    anchor = runs[ri_s]._r
    for segtext, fill in segments:
        if segtext == '':
            continue
        anchor = _insert_like(para, anchor, ref, segtext, fill)
    if suffix:
        _insert_like(para, anchor, ref, suffix, None)


# ── conversion ───────────────────────────────────────────────────────────── #
def convert(docx_bytes, enlx_bytes, do_green=True, do_refmarkers=True,
            highlight='GREEN', apply_near=False, bib_source_bytes=None,
            typed_detect='highlight', ref_styles=('refmark',)):
    """typed_detect in {'highlight','pattern','both'} controls how typed
    citations are found (only relevant when do_green=True).
    ref_styles selects which numbered markers to detect: any of
    'refmark' ([[REF n]]), 'bracket' ([n]), 'paren' ((n))."""
    lib = load_library(enlx_bytes)
    bib, bibtext = {}, {}
    if do_refmarkers and parse_bibliography is not None:
        try:
            bib, bibtext = parse_bibliography(bib_source_bytes or docx_bytes)
        except Exception:
            bib, bibtext = {}, {}

    doc = Document(io.BytesIO(docx_bytes))
    report = []
    use_hl = do_green and typed_detect in ('highlight', 'both')
    use_pat = do_green and typed_detect in ('pattern', 'both')

    def classify(key, surname, year, hint):
        rec, score, kind = match_to_library(surname, year, hint, lib)
        if rec and (kind == 'exact' or (kind == 'near' and apply_near)):
            return 'resolved', (key, rec)
        if rec and kind == 'near':
            return 'suggested', (key, rec)
        return 'missing', (key, None)

    def split_refs(text):
        resolved, suggested, missing = [], [], []
        bucket = {'resolved': resolved, 'suggested': suggested, 'missing': missing}
        for piece in text.split(';'):
            piece = piece.strip()
            if not piece or not re.search(r'\d', piece):
                continue
            sur, year, hint = _parse_ref_text(piece)
            state, item = classify(piece, sur, year, hint)
            bucket[state].append(item)
        return resolved, suggested, missing

    # ---- pass 1: highlighted typed citations ----
    for para in doc.paragraphs:
        runs = para.runs
        n = len(runs)
        i = 0
        while i < n:
            run = runs[i]
            txt = run.text or ''

            if use_hl and _is_green(run, highlight):
                members = [i]; j = i + 1
                while j < n:
                    if _is_green(runs[j], highlight):
                        members.append(j); j += 1
                    elif re.fullmatch(r'[\s;,]*', runs[j].text or '') and j + 1 < n and _is_green(runs[j + 1], highlight):
                        members.append(j); j += 1
                    else:
                        break
                green_text = ''.join(runs[k].text for k in members)
                lead = re.match(r'^([.\s]*)', green_text).group(1)
                body = green_text[len(lead):].lstrip('(').rstrip().rstrip(')').rstrip()
                resolved, suggested, missing = split_refs(body)
                report.append({'kind': 'green', 'orig': green_text.strip(),
                               'resolved': resolved, 'suggested': suggested, 'missing': missing})
                first, last = members[0], members[-1]
                if first - 1 >= 0 and (runs[first - 1].text or '').rstrip().endswith('('):
                    pt = runs[first - 1].text
                    runs[first - 1].text = pt[:pt.rstrip().rfind('(')]
                if last + 1 < n and (runs[last + 1].text or '').lstrip().startswith(')'):
                    ft = runs[last + 1].text; pos = ft.find(')')
                    runs[last + 1].text = ft[:pos] + ft[pos + 1:]
                cite = ('{' + '; '.join(_token(r) for _, r in resolved) + '}') if resolved else ''
                ref_run = runs[members[0]]
                ref_run.text = lead + cite
                ref_run.font.highlight_color = None
                for k in members[1:]:
                    runs[k].text = ''
                    runs[k].font.highlight_color = None
                anchor = ref_run._r
                if suggested:
                    anchor = _insert_like(para, anchor, ref_run,
                        ((' ' if cite else '') + '(' + '; '.join(k for k, _ in suggested) + ')'), ORANGE)
                if missing:
                    anchor = _insert_like(para, anchor, ref_run,
                        ((' ' if (cite or suggested) else '') + '(' + '; '.join(k for k, _ in missing) + ')'), RED)
                i = j
                continue

            i += 1

    # ---- pass 2: pattern-detected typed citations (no highlight) ----
    if use_pat:
        for para in doc.paragraphs:
            for run in list(para.runs):
                t = run.text or ''
                if not t or '[[REF' in t or ('{' in t and '#' in t) or _has_shd(run):
                    continue
                matches = list(_CITE_RE.finditer(t))
                if not matches:
                    continue
                segs = []; pos = 0; touched = False
                for mt in matches:
                    resolved, suggested, missing = split_refs(mt.group(1))
                    if not (resolved or suggested):
                        continue  # not a recognizable citation -> leave it in place
                    touched = True
                    report.append({'kind': 'pattern', 'orig': mt.group(0),
                                   'resolved': resolved, 'suggested': suggested, 'missing': missing})
                    segs.append(('plain', t[pos:mt.start()]))
                    if resolved:
                        segs.append(('plain', '{' + '; '.join(_token(r) for _, r in resolved) + '}'))
                    if suggested:
                        segs.append((ORANGE, (' ' if resolved else '') + '(' + '; '.join(k for k, _ in suggested) + ')'))
                    if missing:
                        segs.append((RED, (' ' if (resolved or suggested) else '') + '(' + '; '.join(k for k, _ in missing) + ')'))
                    pos = mt.end()
                if not touched:
                    continue
                segs.append(('plain', t[pos:]))
                run.text = segs[0][1]
                anchor = run._r
                for fill, segtext in segs[1:]:
                    anchor = _insert_like(para, anchor, run, segtext,
                                          None if fill == 'plain' else fill)

    # ---- pass 3: numbered markers  [[REF n]] / [n] / (n) ----
    if do_refmarkers and ref_styles:
        pats = [(s, _NUM_PAT[s]) for s in ref_styles if s in _NUM_PAT]
        for para in doc.paragraphs:
            full = ''.join(r.text or '' for r in para.runs)
            if not full:
                continue
            found = []
            for style, pat in pats:
                for m in pat.finditer(full):
                    found.append((m.start(), m.end(), style, m.group(1)))
            if not found:
                continue
            found.sort()
            chosen, last_end = [], -1
            for s, e, style, inner in found:
                if s >= last_end:
                    chosen.append((s, e, style, inner)); last_end = e
            actions = []
            for s, e, style, inner in chosen:
                nums = _expand_nums(inner)
                if not nums:
                    continue
                in_bib = any(bib.get(num) is not None for num in nums)
                if style != 'refmark' and not in_bib:
                    continue  # bracketed/parenthesised number that isn't a reference
                resolved, suggested, missing = [], [], []
                bucket = {'resolved': resolved, 'suggested': suggested, 'missing': missing}
                for num in nums:
                    sy = bib.get(num); hint = bibtext.get(num, '')
                    if not sy:
                        missing.append((num, None)); continue
                    st, item = classify(num, sy[0], sy[1], hint)
                    bucket[st].append(item)
                report.append({'kind': style, 'orig': full[s:e],
                               'resolved': resolved, 'suggested': suggested, 'missing': missing})
                cite = ('{' + '; '.join(_token(r) for _, r in resolved) + '}') if resolved else ''
                segs = []
                if cite:
                    segs.append((cite, None))
                if suggested:
                    segs.append(((' ' if cite else '') + _fmt_marker(style, [k for k, _ in suggested]), ORANGE))
                if missing:
                    segs.append(((' ' if (cite or suggested) else '') + _fmt_marker(style, [k for k, _ in missing]), RED))
                actions.append((s, e, segs))
            for s, e, segs in reversed(actions):
                _splice_marker(para, s, e, segs)

    buf = io.BytesIO(); doc.save(buf)
    return _fix_zoom(buf.getvalue()), report


def _fix_zoom(b):
    zin = zipfile.ZipFile(io.BytesIO(b)); out = io.BytesIO()
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zout:
        for it in zin.infolist():
            d = zin.read(it.filename)
            if it.filename == 'word/settings.xml':
                d = re.sub(rb'<w:zoom[^>]*/>', b'<w:zoom w:percent="100"/>', d)
            zout.writestr(it, d)
    return out.getvalue()


# ── summary + report doc ─────────────────────────────────────────────────── #
def summarize(report):
    return {
        'placeholders': len(report),
        'resolved_refs': sum(len(r['resolved']) for r in report),
        'suggested_refs': sum(len(r['suggested']) for r in report),
        'unresolved_refs': sum(len(r['missing']) for r in report),
        'fully_done': sum(1 for r in report if r['resolved'] and not (r['suggested'] or r['missing'])),
    }


def build_report_docx(report, title="Placeholder \u2192 EndNote Conversion Report"):
    ACCENT, FILL = "8B1A1A", "F3ECEC"
    def shade(cell, fill):
        tcPr = cell._tc.get_or_add_tcPr(); sh = OxmlElement('w:shd')
        sh.set(qn('w:val'), 'clear'); sh.set(qn('w:color'), 'auto'); sh.set(qn('w:fill'), fill); tcPr.append(sh)
    def grid(table):
        b = OxmlElement('w:tblBorders')
        for e in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
            x = OxmlElement('w:' + e); x.set(qn('w:val'), 'single'); x.set(qn('w:sz'), '4')
            x.set(qn('w:space'), '0'); x.set(qn('w:color'), 'CCCCCC'); b.append(x)
        tblPr = table._tbl.tblPr; anchor = None
        for c in tblPr:
            if c.tag in (qn('w:shd'), qn('w:tblLayout'), qn('w:tblLook')):
                anchor = c; break
        anchor.addprevious(b) if anchor is not None else tblPr.append(b)

    doc = Document()
    s = doc.sections[0]; s.left_margin = s.right_margin = Inches(0.8); s.top_margin = s.bottom_margin = Inches(0.8)
    doc.styles['Normal'].font.name = 'Calibri'; doc.styles['Normal'].font.size = Pt(10)
    h = doc.add_paragraph(); r = h.add_run(title)
    r.bold = True; r.font.size = Pt(16); r.font.color.rgb = RGBColor.from_string(ACCENT)
    sm = summarize(report)
    sp = doc.add_paragraph(); sr = sp.add_run(
        "%d placeholder(s):  %d reference(s) applied,  %d suggested (orange),  %d unresolved (red)."
        % (sm['placeholders'], sm['resolved_refs'], sm['suggested_refs'], sm['unresolved_refs']))
    sr.font.size = Pt(9); sr.bold = True
    note = doc.add_paragraph(); nr = note.add_run(
        "Applied = inserted as {Author, Year #RecNum}. Suggested (orange in the document) = surname "
        "matched but year/journal differs; verify, then accept or fix. Unresolved (red) = add to library.")
    nr.font.size = Pt(8.5); nr.italic = True; nr.font.color.rgb = RGBColor.from_string("666666")
    doc.add_paragraph()

    table = doc.add_table(rows=1, cols=2); grid(table)
    hdr = table.rows[0]
    for i, t in enumerate(["Placeholder", "Result"]):
        c = hdr.cells[i]; run = c.paragraphs[0].add_run(t)
        run.bold = True; run.font.size = Pt(10); run.font.color.rgb = RGBColor.from_string("FFFFFF"); shade(c, ACCENT)
    trPr = hdr._tr.get_or_add_trPr(); th = OxmlElement('w:tblHeader'); th.set(qn('w:val'), 'true'); trPr.append(th)

    for row in report:
        tr = table.add_row()
        p = tr.cells[0].paragraphs[0]
        run = p.add_run(row['orig'][:90]); run.bold = True; run.font.size = Pt(9.5)
        run.font.color.rgb = RGBColor.from_string(ACCENT); shade(tr.cells[0], FILL)
        cell = tr.cells[1]; cell.paragraphs[0].text = ""; firstpar = [True]
        def line(text, color="000000", bold=False, fill=None):
            para = cell.paragraphs[0] if firstpar[0] else cell.add_paragraph(); firstpar[0] = False
            rr = para.add_run(text); rr.font.size = Pt(9.5); rr.bold = bold
            rr.font.color.rgb = RGBColor.from_string(color)
            if fill:
                _shade_run(rr, fill)
        for key, rec in row['resolved']:
            line("APPLIED   %s  \u2192  %s, %s #%d  (%s)" % (key, rec['sur'], rec['year'], rec['id'], rec['jour'][:24]), "1A6B2A")
        for key, rec in row['suggested']:
            line("SUGGESTED %s  \u2192  %s, %s #%d  (%s)  \u2013 verify" % (key, rec['sur'], rec['year'], rec['id'], rec['jour'][:24]), "7A5A00", fill=ORANGE)
        for key, rec in row['missing']:
            line("UNRESOLVED  %s  \u2013 add to library" % key, "8B1A1A", bold=True, fill=RED)
        if not (row['resolved'] or row['suggested'] or row['missing']):
            line("(no references parsed)", "666666")

    table.autofit = False; table.allow_autofit = False
    for rw in table.rows:
        rw.cells[0].width = Inches(2.0); rw.cells[1].width = Inches(5.0)
    buf = io.BytesIO(); doc.save(buf)
    return _fix_zoom(buf.getvalue())
