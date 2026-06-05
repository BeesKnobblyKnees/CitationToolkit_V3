"""
Bibliography Relink Module
Relinks a text-edited draft whose citations are bare superscript numbers, using
(a) the draft's own numbered bibliography to learn what each number is, and
(b) a citation-intact "old" document as the source of real EndNote field codes.

Matching is by REFERENCE IDENTITY (first author + year), which is immune to the
renumbering that happens when either document is edited. Each in-text citation
group is resolved to its references; if a whole old field code cites exactly that
set it is transplanted intact, otherwise a field code is assembled from the old
document's real per-reference records (correct nested EN.CITE/EN.CITE.DATA form).
Anything unresolved becomes a yellow [[REF n]] placeholder for manual insertion.

Two citation populations are handled:
  * BODY  - bare superscript numbers (rStyle "citsup"), comma-separated by "sup".
  * FOOTNOTES - "References 1, 86, 119..." lists where the numbers carry rStyle
    "crossref" and are separated by plain ", " runs. Only footnotes whose text
    contains "Reference" are touched, so figure/table cross-refs are never altered.
"""

import re
import io
import zipfile
import base64
import html
import docx
from collections import Counter


def _docxml(b):
    with zipfile.ZipFile(io.BytesIO(b)) as z:
        return z.read("word/document.xml").decode("utf-8")


def _surname(author):
    m = re.match(r'^([A-Za-z\-\u00C0-\u017F]+)', author.strip())
    return m.group(1).lower() if m else None


def parse_bibliography(draft_bytes):
    """Map in-text/bibliography number -> (surname, year) and -> display text."""
    d = docx.Document(io.BytesIO(draft_bytes))
    bib, bibtext = {}, {}
    for p in d.paragraphs:
        m = re.match(r'^(\d{1,3})\s+(.*)', p.text.strip())
        if not m:
            continue
        num, rest = int(m.group(1)), m.group(2)
        if len(rest) < 15:
            continue
        yr = re.findall(r'\b(18|19|20)(\d{2})\b', rest)
        am = re.match(r'^([A-Z][a-zA-Z\-]+)', rest)
        bibtext[num] = rest[:110]
        if yr and am:
            bib[num] = (am.group(1).lower(), yr[-1][0] + yr[-1][1])
    return bib, bibtext


def index_old_fieldcodes(old_bytes):
    """Return (fieldxml_by_set, cite_by_id) from the citation-intact old doc."""
    xml = _docxml(old_bytes)
    runs = re.findall(r"<w:r\b(?:(?!</w:r>).)*?</w:r>", xml, re.DOTALL)
    fieldxml_by_set, cite_by_id = {}, {}
    i = 0
    while i < len(runs):
        if 'fldCharType="begin"' in runs[i]:
            depth, blk, j = 0, [], i
            while j < len(runs):
                depth += runs[j].count('fldCharType="begin"') - runs[j].count('fldCharType="end"')
                blk.append(runs[j]); j += 1
                if depth <= 0:
                    break
            block = "".join(blk)
            fd = re.search(r'<w:fldData[^>]*>([\s\S]+?)</w:fldData>', block)
            if fd:
                b64 = "".join(fd.group(1).split()); pad = (4 - len(b64) % 4) % 4
                try:
                    dec = base64.b64decode(b64 + "=" * pad).decode("utf-8", "replace")
                    ids = set()
                    for c in re.findall(r"<Cite>.*?</Cite>", dec, re.DOTALL):
                        au = re.search(r"<Author>([^<]+)</Author>", c)
                        yr = re.search(r"<Year>(\d{4})</Year>", c)
                        if au and yr:
                            s = _surname(au.group(1))
                            if s:
                                cite_by_id.setdefault((s, yr.group(1)), c)
                                ids.add((s, yr.group(1)))
                    if ids:
                        fieldxml_by_set.setdefault(frozenset(ids), block)
                except Exception:
                    pass
            i = j
        else:
            i += 1
    return fieldxml_by_set, cite_by_id


def _strip_disp(c):
    return re.sub(r"<DisplayText>.*?</DisplayText>", "", c, flags=re.DOTALL)


def _build_fieldcode(nums, bib, cite_by_id, rstyle="citsup", superscript=True):
    idents = [bib.get(n) for n in nums]
    if any(x is None or x not in cite_by_id for x in idents):
        return None
    rpr = '<w:rPr><w:rStyle w:val="%s"/></w:rPr>' % rstyle
    face = "superscript" if superscript else "normal"
    cites = []
    for k, ident in enumerate(idents):
        c = cite_by_id[ident]
        if k == 0:
            disp = ('<DisplayText><style face="%s">' % face
                    + ", ".join(str(n) for n in nums) + "</style></DisplayText>")
            c = re.sub(r"(</RecNum>)", r"\1" + disp, _strip_disp(c), count=1)
        else:
            c = _strip_disp(c)
        cites.append(c)
    en = "<EndNote>" + "".join(cites) + "</EndNote>"
    b64 = base64.b64encode(en.encode("utf-8")).decode("ascii")
    show = ", ".join(str(n) for n in nums)
    return (f'<w:r>{rpr}<w:fldChar w:fldCharType="begin"><w:fldData xml:space="preserve">{b64}</w:fldData></w:fldChar></w:r>'
            f'<w:r>{rpr}<w:instrText xml:space="preserve"> ADDIN EN.CITE </w:instrText></w:r>'
            f'<w:r>{rpr}<w:fldChar w:fldCharType="begin"><w:fldData xml:space="preserve">{b64}</w:fldData></w:fldChar></w:r>'
            f'<w:r>{rpr}<w:instrText xml:space="preserve"> ADDIN EN.CITE.DATA </w:instrText></w:r>'
            f'<w:r>{rpr}</w:r>'
            f'<w:r>{rpr}<w:fldChar w:fldCharType="end"/></w:r>'
            f'<w:r>{rpr}<w:fldChar w:fldCharType="separate"/></w:r>'
            f'<w:r>{rpr}<w:t xml:space="preserve">{show}</w:t></w:r>'
            f'<w:r>{rpr}<w:fldChar w:fldCharType="end"/></w:r>')


def _placeholder(nums, rstyle="citsup"):
    return ('<w:r><w:rPr><w:rStyle w:val="%s"/><w:highlight w:val="yellow"/></w:rPr>' % rstyle
            + '<w:t xml:space="preserve">[[REF %s]]</w:t></w:r>' % ",".join(str(n) for n in nums))


def _scan_replace(xml, bib, bibtext, fieldxml_by_set, cite_by_id, construct,
                  cite_style, sep_require_style, build_rstyle, build_superscript):
    """Find consecutive cite-styled number runs, replace each group with a real
    field code (exact transplant / assembled) or a placeholder. Returns
    (new_xml, kinds_counter, placeholders, all_nums)."""
    runs = re.findall(r"<w:r\b(?:(?!</w:r>).)*?</w:r>", xml, re.DOTALL)
    reps, placeholders, all_nums = [], [], []
    cur, cur_runs, last_num = [], [], -1

    def flush():
        nonlocal cur, cur_runs, last_num
        if cur:
            span = "".join(cur_runs[:last_num + 1])          # drop trailing separators
            idents = [bib.get(n) for n in cur]
            fs = frozenset(i for i in idents if i)
            if all(idents) and fs in fieldxml_by_set:
                reps.append((span, fieldxml_by_set[fs], "exact"))
            elif construct and all(idents) and all(i in cite_by_id for i in idents):
                fc = _build_fieldcode(cur, bib, cite_by_id, build_rstyle, build_superscript)
                if fc:
                    reps.append((span, fc, "constructed"))
                else:
                    reps.append((span, _placeholder(cur, build_rstyle), "placeholder"))
                    placeholders.append((cur[:], " | ".join(bibtext.get(n, "?") for n in cur)))
            else:
                reps.append((span, _placeholder(cur, build_rstyle), "placeholder"))
                placeholders.append((cur[:], " | ".join(bibtext.get(n, "?") for n in cur)))
        cur, cur_runs, last_num = [], [], -1

    cite_tok = 'w:val="%s"' % cite_style
    for r in runs:
        tm = re.search(r"<w:t[^>]*>(.*?)</w:t>", r, re.DOTALL)
        txt = html.unescape(tm.group(1)) if tm else ""
        if cite_tok in r and txt.strip().isdigit():
            cur.append(int(txt.strip())); cur_runs.append(r)
            last_num = len(cur_runs) - 1; all_nums.append(int(txt.strip()))
        elif txt.strip() in {",", ";"} and cur and (not sep_require_style or 'w:val="sup"' in r):
            cur_runs.append(r)
        else:
            flush()
    flush()

    out = xml
    for span, repl, _ in reps:
        out = out.replace(span, repl, 1)
    return out, Counter(k for _, _, k in reps), placeholders, all_nums


def _relink_reference_footnotes(foot_xml, bib, bibtext, fieldxml_by_set, cite_by_id, construct):
    """Relink only footnotes whose text mentions 'Reference', leaving all other
    footnotes (figure/table cross-refs, editorial notes) untouched."""
    out = foot_xml
    kinds, phs, nums = Counter(), [], []
    for m in re.finditer(r"<w:footnote\b[^>]*>.*?</w:footnote>", foot_xml, re.DOTALL):
        block = m.group(0)
        vis = "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", block))
        if "Reference" not in vis:
            continue
        nb, k, p, n = _scan_replace(block, bib, bibtext, fieldxml_by_set, cite_by_id,
                                    construct, cite_style="crossref", sep_require_style=False,
                                    build_rstyle="crossref", build_superscript=False)
        if nb != block:
            out = out.replace(block, nb, 1)
            kinds += k
            phs += [(ns, "(footnote) " + tx) for ns, tx in p]
            nums += n
    return out, kinds, phs, nums


def relink(draft_bytes, old_bytes, construct=True):
    """Return (fixed_bytes, report, placeholders[list of (numbers, ref_text)])."""
    bib, bibtext = parse_bibliography(draft_bytes)
    fieldxml_by_set, cite_by_id = index_old_fieldcodes(old_bytes)

    src = zipfile.ZipFile(io.BytesIO(draft_bytes))
    names = set(src.namelist())

    body_xml = src.read("word/document.xml").decode("utf-8")
    body_out, body_kinds, body_ph, body_nums = _scan_replace(
        body_xml, bib, bibtext, fieldxml_by_set, cite_by_id, construct,
        cite_style="citsup", sep_require_style=True,
        build_rstyle="citsup", build_superscript=True)

    foot_out, foot_kinds, foot_ph, foot_nums = None, Counter(), [], []
    if "word/footnotes.xml" in names:
        foot_xml = src.read("word/footnotes.xml").decode("utf-8")
        foot_out, foot_kinds, foot_ph, foot_nums = _relink_reference_footnotes(
            foot_xml, bib, bibtext, fieldxml_by_set, cite_by_id, construct)
        if foot_out == foot_xml:
            foot_out = None  # nothing changed; keep original part verbatim

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for item in src.infolist():
            if item.filename == "word/document.xml":
                z.writestr(item, body_out.encode("utf-8"))
            elif item.filename == "word/footnotes.xml" and foot_out is not None:
                z.writestr(item, foot_out.encode("utf-8"))
            else:
                z.writestr(item, src.read(item.filename))

    cited = set(body_nums) | set(foot_nums)
    report = {
        "locations": sum(body_kinds.values()) + sum(foot_kinds.values()),
        "exact": body_kinds.get("exact", 0) + foot_kinds.get("exact", 0),
        "constructed": body_kinds.get("constructed", 0) + foot_kinds.get("constructed", 0),
        "placeholders": body_kinds.get("placeholder", 0) + foot_kinds.get("placeholder", 0),
        "unique_refs_cited": len(cited),
        "bibliography_entries": len(bibtext),
        "orphan_bib_entries": len(set(bibtext) - cited),
        # body / footnote split (page reads these with .get, safe if absent)
        "body_locations": sum(body_kinds.values()),
        "footnote_locations": sum(foot_kinds.values()),
        "footnote_exact": foot_kinds.get("exact", 0),
        "footnote_constructed": foot_kinds.get("constructed", 0),
        "footnote_placeholders": foot_kinds.get("placeholder", 0),
    }
    return buf.getvalue(), report, body_ph + foot_ph


if __name__ == "__main__":
    draft = open("lld.docx", "rb").read()
    old = open("lld_linked.docx", "rb").read()
    fixed, report, phs = relink(draft, old)
    print("report:", report)
    print("placeholders:", len(phs), "| footnote placeholders:",
          sum(1 for _, t in phs if t.startswith("(footnote)")))
    open("/tmp/_test_relink.docx", "wb").write(fixed)
    x = zipfile.ZipFile(io.BytesIO(fixed)).read("word/document.xml").decode()
    f = zipfile.ZipFile(io.BytesIO(fixed)).read("word/footnotes.xml").decode()
    print("body begins/ends:", x.count('fldCharType="begin"'), x.count('fldCharType="end"'))
    print("foot begins/ends:", f.count('fldCharType="begin"'), f.count('fldCharType="end"'))
