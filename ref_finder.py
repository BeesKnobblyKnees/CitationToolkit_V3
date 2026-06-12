"""
ref_finder.py  --  find references missing from an EndNote library and build
importable EndNote files (.enw / .ris) for them.

Companion to the Placeholder -> EndNote flow (page 11). When a citation
placeholder can't be turned into an EndNote temporary citation because the
reference isn't in the library, feed the reference text here. It queries
CrossRef, returns candidate matches with metadata, and writes .enw / .ris
records you can import into EndNote -- after which the placeholder converter
will resolve them.

No third-party dependencies (stdlib urllib + json). Needs internet access.
CrossRef's public API has no key; pass a contact email (mailto) to use the
faster "polite pool" and be a good citizen.
"""
import json
import time
import urllib.parse
import urllib.request

CROSSREF_URL = "https://api.crossref.org/works"
DEFAULT_TIMEOUT = 20

# CrossRef "type" -> (EndNote %0 reference type, RIS TY type)
_TYPE_MAP = {
    "journal-article":      ("Journal Article", "JOUR"),
    "proceedings-article":  ("Conference Paper", "CPAPER"),
    "book-chapter":         ("Book Section", "CHAP"),
    "book":                 ("Book", "BOOK"),
    "book-part":            ("Book Section", "CHAP"),
    "monograph":            ("Book", "BOOK"),
    "reference-book":       ("Book", "BOOK"),
    "edited-book":          ("Edited Book", "EDBOOK"),
    "report":               ("Report", "RPRT"),
    "dataset":              ("Dataset", "DATA"),
    "dissertation":         ("Thesis", "THES"),
    "posted-content":       ("Journal Article", "JOUR"),  # preprints
}


def _user_agent(mailto=None):
    ua = "CitationToolkit/1.0 (ref_finder)"
    if mailto:
        ua += " (mailto:%s)" % mailto
    return ua


def search_crossref(query, rows=3, mailto=None, timeout=DEFAULT_TIMEOUT):
    """Query CrossRef bibliographic search. Returns a list of raw CrossRef items
    (possibly empty). Network errors propagate to the caller."""
    query = (query or "").strip()
    if not query:
        return []
    params = {
        "query.bibliographic": query,
        "rows": str(max(1, min(int(rows), 10))),
        "select": ("DOI,title,container-title,short-container-title,author,"
                   "issued,published-print,published-online,volume,issue,page,"
                   "ISSN,type,score"),
    }
    if mailto:
        params["mailto"] = mailto
    url = CROSSREF_URL + "?" + urllib.parse.urlencode(params)
    req = urllib.request.Request(url, headers={"User-Agent": _user_agent(mailto)})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        data = json.loads(resp.read().decode("utf-8", "replace"))
    return data.get("message", {}).get("items", []) or []


def _year_from(item):
    # Prefer the print/volume year (what appears in a citation) over the e-pub
    # ("issued") date, which CrossRef often sets to an earlier ahead-of-print day.
    for key in ("published-print", "issued", "published-online"):
        dp = (item.get(key) or {}).get("date-parts") or []
        if dp and dp[0] and dp[0][0]:
            return str(dp[0][0])
    return ""


def crossref_to_fields(item):
    """Normalize a CrossRef item into a flat dict of citation fields."""
    authors = []
    for a in item.get("author", []) or []:
        fam = (a.get("family") or "").strip()
        giv = (a.get("given") or "").strip()
        if fam and giv:
            authors.append("%s, %s" % (fam, giv))
        elif fam:
            authors.append(fam)
        elif a.get("name"):
            authors.append(a["name"].strip())
    title = item["title"][0].strip() if item.get("title") else ""
    ct = item.get("container-title") or item.get("short-container-title") or []
    journal = ct[0].strip() if ct else ""
    issn = item["ISSN"][0].strip() if item.get("ISSN") else ""
    ctype = item.get("type", "")
    enw_type, ris_type = _TYPE_MAP.get(ctype, ("Journal Article", "JOUR"))
    return {
        "type_crossref": ctype,
        "type_enw": enw_type,
        "type_ris": ris_type,
        "authors": authors,
        "title": title,
        "journal": journal,
        "year": _year_from(item),
        "volume": (item.get("volume") or "").strip(),
        "issue": (item.get("issue") or "").strip(),
        "pages": (item.get("page") or "").strip(),
        "doi": (item.get("DOI") or "").strip(),
        "issn": issn,
        "score": item.get("score"),
        "source": "crossref",
    }


def fields_to_enw(f):
    """One EndNote import (.enw / refer) record."""
    lines = ["%0 " + f.get("type_enw", "Journal Article")]
    for a in f.get("authors", []):
        lines.append("%A " + a)
    if f.get("title"):   lines.append("%T " + f["title"])
    if f.get("journal"): lines.append("%J " + f["journal"])
    if f.get("year"):    lines.append("%D " + f["year"])
    if f.get("volume"):  lines.append("%V " + f["volume"])
    if f.get("issue"):   lines.append("%N " + f["issue"])
    if f.get("pages"):   lines.append("%P " + f["pages"])
    if f.get("doi"):     lines.append("%R " + f["doi"])
    if f.get("issn"):    lines.append("%@ " + f["issn"])
    if f.get("doi"):     lines.append("%U https://doi.org/" + f["doi"])
    return "\n".join(lines)


def fields_to_ris(f):
    """One RIS record (EndNote also imports RIS cleanly)."""
    lines = ["TY  - " + f.get("type_ris", "JOUR")]
    for a in f.get("authors", []):
        lines.append("AU  - " + a)
    if f.get("title"):   lines.append("TI  - " + f["title"])
    if f.get("journal"): lines.append("JO  - " + f["journal"])
    if f.get("year"):    lines.append("PY  - " + f["year"])
    if f.get("volume"):  lines.append("VL  - " + f["volume"])
    if f.get("issue"):   lines.append("IS  - " + f["issue"])
    pages = f.get("pages", "")
    if pages and "-" in pages:
        sp, ep = pages.split("-", 1)
        lines.append("SP  - " + sp.strip())
        lines.append("EP  - " + ep.strip())
    elif pages:
        lines.append("SP  - " + pages.strip())
    if f.get("doi"):  lines.append("DO  - " + f["doi"])
    if f.get("issn"): lines.append("SN  - " + f["issn"])
    lines.append("ER  - ")
    return "\n".join(lines)


def build_enw(fields_list):
    """Combined .enw text for a list of field dicts (blank line between records)."""
    return "\n\n".join(fields_to_enw(f) for f in fields_list) + "\n"


def build_ris(fields_list):
    return "\n".join(fields_to_ris(f) for f in fields_list) + "\n"


def build_report_csv(rows):
    """Build a CSV string from report rows. Each row is a dict with any of:
    Input, Status, Source, Authors, Title, Journal, Year, Volume, Issue, Pages,
    DOI, PMID, Included."""
    import csv, io as _io
    cols = ["Input", "Status", "Source", "Authors", "Title", "Journal",
            "Year", "Volume", "Issue", "Pages", "DOI", "PMID", "Included"]
    buf = _io.StringIO()
    w = csv.DictWriter(buf, fieldnames=cols, extrasaction="ignore")
    w.writeheader()
    for r in rows:
        w.writerow(r)
    return buf.getvalue()


# ── PubMed fallback (NCBI E-utilities) ──────────────────────────────────────
EUTILS = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils"


def _eutils_params(extra, api_key, email):
    p = dict(extra)
    p["tool"] = "CitationToolkit"
    if email:
        p["email"] = email
    if api_key:
        p["api_key"] = api_key
    return p


def esearch_pubmed(query, retmax=3, api_key=None, email=None, timeout=DEFAULT_TIMEOUT):
    """Return a list of PMIDs for a free-text query."""
    query = (query or "").strip()
    if not query:
        return []
    params = _eutils_params(
        {"db": "pubmed", "term": query, "retmode": "json",
         "retmax": str(max(1, min(int(retmax), 10)))}, api_key, email)
    url = EUTILS + "/esearch.fcgi?" + urllib.parse.urlencode(params)
    req = urllib.request.Request(url, headers={"User-Agent": _user_agent(email)})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        data = json.loads(resp.read().decode("utf-8", "replace"))
    return data.get("esearchresult", {}).get("idlist", []) or []


def esummary_pubmed(pmids, api_key=None, email=None, timeout=DEFAULT_TIMEOUT):
    """Return raw esummary docsum dicts (in PMID order) for a list of PMIDs."""
    pmids = [str(p) for p in pmids if str(p).strip()]
    if not pmids:
        return []
    params = _eutils_params(
        {"db": "pubmed", "id": ",".join(pmids), "retmode": "json"}, api_key, email)
    url = EUTILS + "/esummary.fcgi?" + urllib.parse.urlencode(params)
    req = urllib.request.Request(url, headers={"User-Agent": _user_agent(email)})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        data = json.loads(resp.read().decode("utf-8", "replace"))
    result = data.get("result", {}) or {}
    return [result[u] for u in result.get("uids", []) if u in result]


def _pubmed_author(name):
    """PubMed gives 'Family II' (e.g. 'Al Ramlawi A'); emit 'Family, II'."""
    name = (name or "").strip()
    if not name or " " not in name:
        return name
    head, tail = name.rsplit(" ", 1)
    # tail is initials when it's short and all upper-case letters
    if tail and tail.isalpha() and tail.isupper() and len(tail) <= 4:
        return "%s, %s" % (head, tail)
    return name


def _pubmed_year(doc):
    for key in ("pubdate", "epubdate", "sortpubdate"):
        v = (doc.get(key) or "").strip()
        if v:
            for tok in v.replace("/", " ").split():
                if len(tok) == 4 and tok.isdigit():
                    return tok
    return ""


def pubmed_to_fields(doc):
    """Normalize a PubMed esummary docsum into the same flat field dict."""
    authors = [_pubmed_author(a.get("name", "")) for a in doc.get("authors", [])
               if a.get("name")]
    doi = ""
    for aid in doc.get("articleids", []):
        if aid.get("idtype") == "doi" and aid.get("value"):
            doi = aid["value"].strip()
            break
    journal = (doc.get("fulljournalname") or doc.get("source") or "").strip()
    issn = (doc.get("issn") or doc.get("essn") or "").strip()
    return {
        "type_crossref": "journal-article",
        "type_enw": "Journal Article",
        "type_ris": "JOUR",
        "authors": authors,
        "title": (doc.get("title") or "").strip().rstrip("."),
        "journal": journal,
        "year": _pubmed_year(doc),
        "volume": (doc.get("volume") or "").strip(),
        "issue": (doc.get("issue") or "").strip(),
        "pages": (doc.get("pages") or "").strip(),
        "doi": doi,
        "issn": issn,
        "score": None,          # PubMed esummary has no relevance score
        "source": "pubmed",
        "pmid": str(doc.get("uid") or ""),
    }


def search_pubmed(query, rows=3, api_key=None, email=None, timeout=DEFAULT_TIMEOUT):
    """One-call convenience: esearch -> esummary -> field dicts."""
    pmids = esearch_pubmed(query, retmax=rows, api_key=api_key, email=email, timeout=timeout)
    if not pmids:
        return []
    docs = esummary_pubmed(pmids, api_key=api_key, email=email, timeout=timeout)
    return [pubmed_to_fields(d) for d in docs]


def find_references(ref_strings, rows=3, mailto=None, use_pubmed=True,
                    pubmed_api_key=None, pause=0.2, timeout=DEFAULT_TIMEOUT):
    """For each reference string, query CrossRef; if CrossRef returns nothing,
    fall back to PubMed (when use_pubmed). Returns a list of result dicts:
        {query, status, source, candidates:[fields,...], error}
    status is one of: 'found' (>=1 candidate), 'notfound', 'error'.
    'source' notes where the candidates came from ('crossref' | 'pubmed' | '')."""
    results = []
    for q in ref_strings:
        q = (q or "").strip()
        if not q:
            continue
        rec = {"query": q, "status": "notfound", "source": "",
               "candidates": [], "error": None}
        errs = []
        # 1) CrossRef
        try:
            items = search_crossref(q, rows=rows, mailto=mailto, timeout=timeout)
            cands = [crossref_to_fields(it) for it in items]
            if cands:
                rec["candidates"] = cands
                rec["status"] = "found"
                rec["source"] = "crossref"
        except Exception as e:
            errs.append("CrossRef: %s" % e)
        # 2) PubMed fallback (only if CrossRef found nothing)
        if not rec["candidates"] and use_pubmed:
            if pause:
                time.sleep(pause)
            try:
                cands = search_pubmed(q, rows=rows, api_key=pubmed_api_key,
                                      email=mailto, timeout=timeout)
                if cands:
                    rec["candidates"] = cands
                    rec["status"] = "found"
                    rec["source"] = "pubmed"
            except Exception as e:
                errs.append("PubMed: %s" % e)
        if not rec["candidates"]:
            rec["status"] = "error" if errs else "notfound"
        if errs:
            rec["error"] = "; ".join(errs)
        results.append(rec)
        if pause:
            time.sleep(pause)
    return results
