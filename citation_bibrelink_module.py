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
"""

import re
import io
import zipfile
import base64
import html
import docx

_RPR = '<w:rPr><w:rStyle w:val="citsup"/></w:rPr>'


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


def _build_fieldcode(nums, bib, cite_by_id):
    idents = [bib.get(n) for n in nums]
    if any(x is None or x not in cite_by_id for x in idents):
        return None
    cites = []
    for k, ident in enumerate(idents):
        c = cite_by_id[ident]
        if k == 0:
            disp = '<DisplayText><style face="superscript">' + ", ".join(str(n) for n in nums) + "</style></DisplayText>"
            c = re.sub(r"(</RecNum>)", r"\1" + disp, _strip_disp(c), count=1)
        else:
            c = _strip_disp(c)
        cites.append(c)
    en = "<EndNote>" + "".join(cites) + "</EndNote>"
    b64 = base64.b64encode(en.encode("utf-8")).decode("ascii")
    show = ", ".join(str(n) for n in nums)
    return (f'<w:r>{_RPR}<w:fldChar w:fldCharType="begin"><w:fldData xml:space="preserve">{b64}</w:fldData></w:fldChar></w:r>'
            f'<w:r>{_RPR}<w:instrText xml:space="preserve"> ADDIN EN.CITE </w:instrText></w:r>'
            f'<w:r>{_RPR}<w:fldChar w:fldCharType="begin"><w:fldData xml:space="preserve">{b64}</w:fldData></w:fldChar></w:r>'
            f'<w:r>{_RPR}<w:instrText xml:space="preserve"> ADDIN EN.CITE.DATA </w:instrText></w:r>'
            f'<w:r>{_RPR}</w:r>'
            f'<w:r>{_RPR}<w:fldChar w:fldCharType="end"/></w:r>'
            f'<w:r>{_RPR}<w:fldChar w:fldCharType="separate"/></w:r>'
            f'<w:r>{_RPR}<w:t xml:space="preserve">{show}</w:t></w:r>'
            f'<w:r>{_RPR}<w:fldChar w:fldCharType="end"/></w:r>')


def _placeholder(nums):
    return ('<w:r><w:rPr><w:rStyle w:val="citsup"/><w:highlight w:val="yellow"/></w:rPr>'
            f'<w:t xml:space="preserve">[[REF {",".join(str(n) for n in nums)}]]</w:t></w:r>')


def relink(draft_bytes, old_bytes, construct=True):
    """Return (fixed_bytes, report, placeholders[list of (numbers, ref_text)])."""
    bib, bibtext = parse_bibliography(draft_bytes)
    fieldxml_by_set, cite_by_id = index_old_fieldcodes(old_bytes)

    raw = _docxml(draft_bytes)
    runs = re.findall(r"<w:r\b(?:(?!</w:r>).)*?</w:r>", raw, re.DOTALL)

    cur, cur_runs, reps, placeholders = [], [], [], []
    all_nums = []

    def flush():
        if not cur:
            return
        span = "".join(cur_runs)
        idents = [bib.get(n) for n in cur]
        fs = frozenset(i for i in idents if i)
        if all(idents) and fs in fieldxml_by_set:
            reps.append((span, fieldxml_by_set[fs], "exact"))
        elif construct and all(idents) and all(i in cite_by_id for i in idents):
            fc = _build_fieldcode(cur, bib, cite_by_id)
            if fc:
                reps.append((span, fc, "constructed")); return
            reps.append((span, _placeholder(cur), "placeholder"))
            placeholders.append((cur[:], " | ".join(bibtext.get(n, "?") for n in cur)))
        else:
            reps.append((span, _placeholder(cur), "placeholder"))
            placeholders.append((cur[:], " | ".join(bibtext.get(n, "?") for n in cur)))

    for r in runs:
        tm = re.search(r"<w:t[^>]*>(.*?)</w:t>", r, re.DOTALL)
        txt = html.unescape(tm.group(1)) if tm else ""
        if 'w:val="citsup"' in r and txt.strip().isdigit():
            cur.append(int(txt.strip())); cur_runs.append(r); all_nums.append(int(txt.strip()))
        elif 'w:val="sup"' in r and txt.strip() in {",", ";"} and cur:
            cur_runs.append(r)
        else:
            flush(); cur, cur_runs = [], []
    flush()

    from collections import Counter
    kinds = Counter(k for _, _, k in reps)
    out = raw
    for span, repl, _ in reps:
        out = out.replace(span, repl, 1)

    src = zipfile.ZipFile(io.BytesIO(draft_bytes)); buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for item in src.infolist():
            z.writestr(item, out.encode("utf-8") if item.filename == "word/document.xml"
                       else src.read(item.filename))

    report = {
        "locations": len(reps),
        "exact": kinds.get("exact", 0),
        "constructed": kinds.get("constructed", 0),
        "placeholders": kinds.get("placeholder", 0),
        "unique_refs_cited": len(set(all_nums)),
        "bibliography_entries": len(bibtext),
        "orphan_bib_entries": len(set(bibtext) - set(all_nums)),
    }
    return buf.getvalue(), report, placeholders


if __name__ == "__main__":
    draft = open("lld.docx", "rb").read()
    old = open("lld_linked.docx", "rb").read()
    fixed, report, phs = relink(draft, old)
    print("report:", report)
    print("placeholders:", len(phs))
    open("/tmp/_test_relink.docx", "wb").write(fixed)
    # structure check
    x = zipfile.ZipFile(io.BytesIO(fixed)).read("word/document.xml").decode()
    print("begins/ends:", x.count('fldCharType="begin"'), x.count('fldCharType="end"'))
