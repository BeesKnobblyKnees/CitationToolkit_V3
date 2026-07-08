"""Chapter PDF page-range generator (TOC + pdftotext).

get_chapters(pdf_path) -> chapters from PDF bookmarks (instant).
chapter_ranges(pdf_path, num) -> total / text / bibliography ranges in PDF pages
   AND printed book pages. Scans only the chapter's pages via pdftotext (fast).
"""
import fitz, re, subprocess

def get_chapters(pdf_path):
    doc=fitz.open(pdf_path); toc=doc.get_toc(); doc.close()
    seen={}
    for lvl,title,pg in toc:
        m=re.match(r'\s*(\d+)\s*[-\u2010\u2013\u2014]\s*(.+)', title)
        if m and lvl<=4:
            num=int(m.group(1))
            if num not in seen: seen[num]={"num":num,"title":re.sub(r'\u00ad','',m.group(2)).strip(),"pdf_start":pg}
    return sorted(seen.values(), key=lambda x:x["pdf_start"])

def _labels(pdf_path,start,end):
    out=subprocess.run(['pdftotext','-f',str(start),'-l',str(end),pdf_path,'-'],
                       capture_output=True,text=True,errors='ignore').stdout
    recs=[]
    for off,pg in enumerate(out.split('\f')):
        lines=[l.strip() for l in pg.splitlines() if l.strip()][:2]
        h=" ".join(lines)
        h=re.sub(r'CHAPTER\s+\d+','',h)                    # drop chapter number
        h=re.sub(r'(?:Fig(?:ure)?|Table|Box)\.?\s*[\d.\-eE]+','',h, flags=re.I)  # drop fig/table refs
        h=re.sub(r'\b\d+\s*[-\u2010\u2013]\s*\d+\b','',h)  # drop "34-140" style figure numbers
        m=re.search(r'(\d{1,4}\.e\d+)', h)                 # online reference page
        if m: pr,is_ref=m.group(1),True
        else:
            nums=re.findall(r'\b(\d{3,4})\b', h)           # printed body page (3-4 digits)
            pr,is_ref=(nums[0] if nums else None),False
        recs.append({"pdf":start+off,"printed":pr,"is_ref":is_ref})
    return recs

def _num(lbl):
    m=re.match(r'(\d+)', lbl or ''); return int(m.group(1)) if m else None

def chapter_ranges(pdf_path, chapter_num):
    chs=get_chapters(pdf_path)
    ch=next((c for c in chs if c["num"]==chapter_num), None)
    if not ch: return None
    nxt=next((c for c in chs if c["pdf_start"]>ch["pdf_start"]), None)
    doc=fitz.open(pdf_path); npg=doc.page_count; doc.close()
    start=ch["pdf_start"]; end=(nxt["pdf_start"]-1) if nxt else npg
    recs=_labels(pdf_path,start,end)
    while recs and recs[-1]["printed"] is None and not recs[-1]["is_ref"]:
        recs.pop()                              # trim trailing divider/blank pages
    body=[r for r in recs if not r["is_ref"]]; refs=[r for r in recs if r["is_ref"]]
    def rng(ps): return (ps[0]["pdf"],ps[-1]["pdf"]) if ps else None

    # reliable printed labels: the ".eN" base equals the last text page number
    base=None
    for r in refs:
        m=re.match(r'(\d+)\.e\d+', r["printed"] or ''); 
        if m: base=m.group(1); break
    # text printed start = first plausible body label; end = base (if refs) else last body label
    body_nums=[_num(r["printed"]) for r in body if r["printed"]]
    text_start = next((r["printed"] for r in body if r["printed"]), None)
    if base:
        text_end=base
        bib_start, bib_end = f"{base}.e1", f"{base}.e{len(refs)}"
    else:
        text_end = body[-1]["printed"] if body else None
        bib_start=bib_end=None
    # guard: if extracted text_start is an outlier vs text_end, fall back to count-based
    if base and text_start and _num(text_start) and _num(text_start) > int(base):
        text_start = str(int(base) - max(len(body)-1,0))
    return {"chapter":chapter_num,"title":ch["title"],
        "total":{"pdf":rng(recs),"printed":(text_start, bib_end or text_end),"pages":len(recs)},
        "text":{"pdf":rng(body),"printed":(text_start, text_end),"pages":len(body)},
        "bibliography":{"pdf":rng(refs),"printed":(bib_start, bib_end),"pages":len(refs)}}

if __name__=="__main__":
    import json,time
    for n in (30,29,20):
        t=time.time(); r=chapter_ranges("full2.pdf",n)
        print("ch%d (%.1fs):"%(n,time.time()-t), json.dumps(r))
