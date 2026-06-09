"""
placeholder_convert.py  --  turn citation placeholders into EndNote temporary
citations ({Author, Year #RecNum}) by resolving them against an .enlx library.

Two placeholder kinds, both handled in one pass:
  * Green/typed author-year citations  (e.g. "(Hodgen JT 1920 Arch Surg)")
        -> matched to the library by author + year (+ journal hint).
  * [[REF n, n, n]] markers from Bibliography Relink
        -> each number resolved via the document's numbered bibliography
           (number -> author/year) and then matched to the library.

Output is ready for EndNote: open in Word with the SAME library selected and run
Update Citations and Bibliography (or feed it to citation_rebuild). The produced
{Author, Year #RecNum} record numbers are those of the supplied .enlx.

Confident (exact-year) matches are applied automatically. Matches where only the
surname lines up (different year/journal) are reported as SUGGESTIONS and are not
applied unless apply_near=True. Anything unresolved is left in place and flagged.
"""
import io, re, zipfile, sqlite3, difflib, html
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

try:
    from citation_bibrelink_module import parse_bibliography
except Exception:
    parse_bibliography = None


# ── library ────────────────────────────────────────────────────────────────
def load_library(enlx_bytes):
    """Read refs (id, author, year, title, journal) from an .enlx (zip w/ SQLite)."""
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
    """Return (record|None, score, kind) where kind in {'exact','near','none'}."""
    sl = (surname or '').lower()
    jh = (journal_hint or '')
    jtoks = [t.lower() for t in re.findall(r'[A-Z][a-z]{2,}', jh)]
    def jbonus(r):
        return 0.3 if jtoks and any(t in r['jour'].lower() for t in jtoks) else 0.0
    # exact year
    best = None
    for r in lib:
        if r['year'] != year or not year:
            continue
        ratio = difflib.SequenceMatcher(None, sl, r['surl']).ratio()
        starts = bool(sl) and (r['surl'].startswith(sl[:4]) or sl.startswith(r['surl'][:4]))
        score = ratio + (0.25 if starts else 0) + jbonus(r)
        if score >= 0.85 and (best is None or score > best[0]):
            best = (score, r)
    if best:
        return best[1], best[0], 'exact'
    # near: surname matches strongly, any year
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
    """Green typed ref -> (surname, year, journal_hint)."""
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


# ── conversion ───────────────────────────────────────────────────────────── #
def convert(docx_bytes, enlx_bytes, do_green=True, do_refmarkers=True,
            highlight='GREEN', apply_near=False, bib_source_bytes=None):
    lib = load_library(enlx_bytes)
    bib = {}
    bibtext = {}
    if do_refmarkers and parse_bibliography is not None:
        try:
            bib, bibtext = parse_bibliography(bib_source_bytes or docx_bytes)
        except Exception:
            bib, bibtext = {}, {}

    doc = Document(io.BytesIO(docx_bytes))
    report = []

    def resolve(surname, year, hint):
        rec, score, kind = match_to_library(surname, year, hint, lib)
        return rec, kind

    for para in doc.paragraphs:
        runs = para.runs
        n = len(runs)
        i = 0
        while i < n:
            run = runs[i]
            txt = run.text or ''

            # ---- [[REF n, n, ...]] marker (single run) ----
            if do_refmarkers and '[[REF' in txt:
                m = re.search(r'\[\[REF\s*([\d,\s]+)\]\]', txt)
                if m:
                    nums = [int(x) for x in re.findall(r'\d+', m.group(1))]
                    resolved, near, unresolved = [], [], []
                    for num in nums:
                        sy = bib.get(num)
                        hint = bibtext.get(num, '')
                        if not sy:
                            unresolved.append((num, None)); continue
                        rec, kind = resolve(sy[0], sy[1], hint)
                        if rec and (kind == 'exact' or (kind == 'near' and apply_near)):
                            resolved.append((num, rec))
                        elif rec and kind == 'near':
                            near.append((num, rec)); unresolved.append((num, rec))
                        else:
                            unresolved.append((num, None))
                    report.append({'kind': 'refmarker', 'orig': m.group(0),
                                   'resolved': resolved, 'near': near, 'unresolved': unresolved})
                    if resolved:
                        cite = '{' + '; '.join(_token(r) for _, r in resolved) + '}'
                        leftover = [num for num, _ in unresolved]
                        tail = (' [[REF %s]]' % ', '.join(str(x) for x in leftover)) if leftover else ''
                        run.text = txt[:m.start()] + cite + tail + txt[m.end():]
                        if not leftover:
                            run.font.highlight_color = None
                    i += 1
                    continue

            # ---- green typed citation cluster ----
            if do_green and _is_green(run, highlight):
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
                resolved, near, unresolved = [], [], []
                for piece in body.split(';'):
                    piece = piece.strip()
                    if not piece or not re.search(r'\d', piece):
                        continue
                    sur, year, hint = _parse_ref_text(piece)
                    rec, kind = resolve(sur, year, hint)
                    if rec and (kind == 'exact' or (kind == 'near' and apply_near)):
                        resolved.append((piece, rec))
                    elif rec and kind == 'near':
                        near.append((piece, rec)); unresolved.append((piece, rec))
                    else:
                        unresolved.append((piece, None))
                report.append({'kind': 'green', 'orig': green_text.strip(),
                               'resolved': resolved, 'near': near, 'unresolved': unresolved})
                if resolved:
                    cite = '{' + '; '.join(_token(r) for _, r in resolved) + '}'
                    leftover_txt = '; '.join(p for p, _ in unresolved)
                    repl = lead + cite + ((' (' + leftover_txt + ')') if leftover_txt else '')
                    # strip wrapping parens from neighbours
                    first = members[0]
                    if first - 1 >= 0 and (runs[first - 1].text or '').rstrip().endswith('('):
                        pt = runs[first - 1].text
                        runs[first - 1].text = pt[:pt.rstrip().rfind('(')]
                    last = members[-1]
                    if last + 1 < n and (runs[last + 1].text or '').lstrip().startswith(')'):
                        ft = runs[last + 1].text; pos = ft.find(')')
                        runs[last + 1].text = ft[:pos] + ft[pos + 1:]
                    runs[members[0]].text = repl
                    runs[members[0]].font.highlight_color = None
                    for k in members[1:]:
                        runs[k].text = ''
                        runs[k].font.highlight_color = None
                i = j
                continue

            i += 1

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
        'suggested_refs': sum(len(r['near']) for r in report),
        'unresolved_refs': sum(len([u for u in r['unresolved'] if u[1] is None]) for r in report),
        'fully_done': sum(1 for r in report if r['resolved'] and not r['unresolved']),
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
        "%d placeholder(s):  %d reference(s) applied,  %d suggested (not applied),  %d unresolved."
        % (sm['placeholders'], sm['resolved_refs'], sm['suggested_refs'], sm['unresolved_refs']))
    sr.font.size = Pt(9); sr.bold = True
    note = doc.add_paragraph(); nr = note.add_run(
        "Applied = inserted as {Author, Year #RecNum}. Suggested = surname matched but year/journal "
        "differs; verify and apply manually (or re-run with 'apply close matches'). Unresolved = add to library.")
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
        cell = tr.cells[1]; cell.paragraphs[0].text = ""; firstpar = True
        def line(label, color="000000", bold=False):
            nonlocal firstpar
            para = cell.paragraphs[0] if firstpar else cell.add_paragraph(); firstpar = False
            rr = para.add_run(label); rr.font.size = Pt(9.5); rr.bold = bold
            rr.font.color.rgb = RGBColor.from_string(color)
        for key, rec in row['resolved']:
            line("APPLIED  %s  ->  %s, %s #%d  (%s)" % (key, rec['sur'], rec['year'], rec['id'], rec['jour'][:24]), "1A6B2A")
        for key, rec in row['near']:
            line("SUGGEST  %s  ->  %s, %s #%d  (%s)  [verify]" % (key, rec['sur'], rec['year'], rec['id'], rec['jour'][:24]), "8A6D00")
        for key, rec in row['unresolved']:
            if rec is None:
                line("UNRESOLVED  %s  ->  add to library" % key, "8B1A1A", bold=True)
        if not (row['resolved'] or row['near'] or row['unresolved']):
            line("(no references parsed)", "666666")

    table.autofit = False; table.allow_autofit = False
    for rw in table.rows:
        rw.cells[0].width = Inches(2.0); rw.cells[1].width = Inches(5.0)
    buf = io.BytesIO(); doc.save(buf)
    return _fix_zoom(buf.getvalue())
