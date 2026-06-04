"""
Citation Verifier Module — three confidence tiers, all keyed to the text that
PRECEDES each citation (the claim a superscript supports).

Tier A  topical screen : preceding claim vs reference title + MeSH keywords.
Tier B  published cross-check : is the preceding claim retained published prose?
        (matches the claim against the published chapter text — robust to PDF
        superscript loss; confirms the citation sits on original, vetted content.)
Tier C  PubMed abstract : fetch the reference's abstract from PubMed E-utilities
        and score its overlap with the claim. Needs internet (runs on the user's
        machine). Optional.

Output is a per-citation review table, NOT a correctness certificate: it tells
the editor which citations are confirmed, which are merely plausible, and which
look off and deserve a human read.
"""

import re
import io
import zipfile
import base64
import html
import time
import urllib.request
import urllib.parse
from difflib import SequenceMatcher

_STOP = set("the a an of and or to in for with on at by from is are was were be as that this "
            "these those it its their his her not no can may might also been being have has had "
            "which who whom into than then thus such between among during after before over under "
            "but if when while where each per due via about within without".split())


def _docxml(b):
    with zipfile.ZipFile(io.BytesIO(b)) as z:
        return z.read("word/document.xml").decode("utf-8")


def _terms(text):
    return {w for w in re.findall(r"[a-z]{4,}", text.lower()) if w not in _STOP}


def extract_citations(relinked_bytes):
    """
    Walk the relinked doc. For each EndNote field code, return:
      {claim: preceding text, refs: [{author,year,title,keywords}], display: nums}
    Claim = text since the previous citation/sentence boundary, up to this cite.
    """
    xml = _docxml(relinked_bytes)
    runs = re.findall(r"<w:r\b(?:(?!</w:r>).)*?</w:r>", xml, re.DOTALL)
    cites = []
    plain = ""
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
            refs = []
            if fd:
                b64 = "".join(fd.group(1).split()); pad = (4 - len(b64) % 4) % 4
                try:
                    dec = base64.b64decode(b64 + "=" * pad).decode("utf-8", "replace")
                    for c in re.findall(r"<Cite>.*?</Cite>", dec, re.DOTALL):
                        au = re.search(r"<Author>([^<]+)</Author>", c)
                        yr = re.search(r"<Year>(\d{4})</Year>", c)
                        ti = re.search(r"<title>(.*?)</title>", c, re.DOTALL)
                        kw = re.findall(r"<keyword>(.*?)</keyword>", c, re.DOTALL)
                        refs.append({
                            "author": au.group(1).strip() if au else "?",
                            "year": yr.group(1) if yr else "?",
                            "title": re.sub(r"<[^>]+>", "", ti.group(1)).strip() if ti else "",
                            "keywords": " ".join(re.sub(r"<[^>]+>", "", k) for k in kw),
                        })
                except Exception:
                    pass
            # claim = last sentence(s) of accumulated plain text
            claim = plain.strip()
            mlast = re.split(r"(?<=[.;:])\s", claim)
            claim_sent = " ".join(mlast[-2:]) if len(mlast) > 1 else claim
            cites.append({"claim": claim_sent[-300:], "refs": refs})
            plain = ""
            i = j
        else:
            tm = re.search(r"<w:t[^>]*>(.*?)</w:t>", runs[i], re.DOTALL)
            if tm:
                plain += html.unescape(tm.group(1))
            i += 1
    return cites


# ---- Tier A: topical overlap ----
def tier_a(claim, ref):
    ct = _terms(claim)
    rt = _terms(ref["title"] + " " + ref["keywords"])
    if not ct or not rt:
        return 0.0
    return len(ct & rt) / max(len(ct), 1)


# ---- Tier B: claim retained in published chapter ----
def extract_published_text(pdf_bytes, first_page=955, last_page=1015):
    import pdfplumber
    out = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        lo, hi = max(0, first_page), min(len(pdf.pages), last_page)
        for p in pdf.pages[lo:hi]:
            t = p.extract_text() or ""
            if "Limb Length Discrepancy" in t or out:
                out.append(t)
    return re.sub(r"\s+", " ", " ".join(out)).lower()


def tier_b(claim, published_text):
    """Best local match ratio of the claim within the published chapter text."""
    c = re.sub(r"\s+", " ", claim.lower()).strip()
    c = re.sub(r"\d+", "", c)  # drop citation digits
    if len(c) < 25 or not published_text:
        return 0.0
    # slide a window the size of the claim and take the best SequenceMatcher ratio
    best = 0.0
    step = max(20, len(c) // 2)
    for k in range(0, len(published_text) - len(c), step):
        seg = published_text[k:k + len(c) + 40]
        r = SequenceMatcher(None, c, seg).ratio()
        if r > best:
            best = r
            if best > 0.85:
                break
    return best


# ---- Tier C: PubMed abstract ----
def pubmed_abstract(ref, tool="CitationToolkit", email=""):
    """Search PubMed for the reference and return its abstract text (or '')."""
    base = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/"
    au = ref["author"].split(",")[0].split()[0] if ref["author"] not in ("", "?") else ""
    term = f'{au}[Author] AND {ref["year"]}[PDAT]'
    tw = [w for w in re.findall(r"[A-Za-z]{5,}", ref["title"])][:4]
    if tw:
        term += " AND " + " AND ".join(f"{w}[Title]" for w in tw)
    q = urllib.parse.urlencode({"db": "pubmed", "term": term, "retmax": "1",
                                "tool": tool, "email": email})
    hdr = {"User-Agent": f"{tool} (mailto:{email})"}
    try:
        req = urllib.request.Request(base + "esearch.fcgi?" + q, headers=hdr)
        xml = urllib.request.urlopen(req, timeout=12).read().decode("utf-8", "replace")
        pmid = re.search(r"<Id>(\d+)</Id>", xml)
        if not pmid:
            return ""
        time.sleep(0.34)
        q2 = urllib.parse.urlencode({"db": "pubmed", "id": pmid.group(1),
                                     "rettype": "abstract", "retmode": "xml",
                                     "tool": tool, "email": email})
        req2 = urllib.request.Request(base + "efetch.fcgi?" + q2, headers=hdr)
        ab = urllib.request.urlopen(req2, timeout=12).read().decode("utf-8", "replace")
        texts = re.findall(r"<AbstractText[^>]*>(.*?)</AbstractText>", ab, re.DOTALL)
        return re.sub(r"<[^>]+>", " ", " ".join(texts))
    except Exception:
        return ""


def tier_c(claim, abstract):
    if not abstract:
        return None
    ct, at = _terms(claim), _terms(abstract)
    if not ct or not at:
        return 0.0
    return len(ct & at) / max(len(ct), 1)


def verify(relinked_bytes, published_pdf_bytes=None, use_pubmed=False,
           pubmed_email="", progress=None):
    cites = extract_citations(relinked_bytes)
    published = extract_published_text(published_pdf_bytes) if published_pdf_bytes else ""
    rows = []
    for idx, cu in enumerate(cites):
        if progress:
            progress(idx, len(cites))
        claim = cu["claim"]
        ref0 = cu["refs"][0] if cu["refs"] else {"author": "?", "year": "?", "title": "", "keywords": ""}
        a = max((tier_a(claim, r) for r in cu["refs"]), default=0.0)
        b = tier_b(claim, published) if published else None
        c = None
        if use_pubmed and cu["refs"]:
            ab = pubmed_abstract(ref0, email=pubmed_email)
            c = tier_c(claim, ab)
            time.sleep(0.34)
        # verdict
        signals = [s for s in (a, c) if s is not None]
        topical = max(signals) if signals else 0.0
        if b is not None and b > 0.7 and topical >= 0.12:
            verdict = "verified"          # retained published prose + on-topic ref
        elif topical >= 0.25 or (c is not None and c >= 0.2):
            verdict = "supported"
        elif topical >= 0.1 or (b is not None and b > 0.7):
            verdict = "plausible"
        else:
            verdict = "review"
        rows.append({
            "claim": claim[-120:],
            "reference": f'{ref0["author"]} ({ref0["year"]}) {ref0["title"][:60]}',
            "topical_A": round(a, 2),
            "published_B": (round(b, 2) if b is not None else None),
            "pubmed_C": (round(c, 2) if c is not None else None),
            "verdict": verdict,
        })
    return rows


if __name__ == "__main__":
    relinked = open("/mnt/user-data/outputs/LLD_FIXED_v3.docx", "rb").read()
    pdf = open("/mnt/user-data/uploads/Tachdjian_s_Full_Text.pdf", "rb").read()
    rows = verify(relinked, pdf, use_pubmed=False)
    from collections import Counter
    print("citations checked:", len(rows))
    print("verdicts:", dict(Counter(r["verdict"] for r in rows)))
    for r in rows[:6]:
        print(f'  [{r["verdict"]:9}] A={r["topical_A"]} B={r["published_B"]} | {r["reference"][:50]}')
        print(f'             claim: …{r["claim"][-70:]}')
