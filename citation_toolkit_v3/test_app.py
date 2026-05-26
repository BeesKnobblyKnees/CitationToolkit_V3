"""
Citation Toolkit v3 — automated test suite.
Run from inside the citation_toolkit_v3 folder:
    python3 test_app.py
All tests must pass before uploading to GitHub.
"""
import sys, re, ast, io, zipfile, base64, os
from pathlib import Path

PASS = "\033[92m✓\033[0m"
FAIL = "\033[91m✗\033[0m"
results = []

BASE   = Path(__file__).parent
BLOCKS = BASE / "_app_blocks"

def test(name, fn):
    try:
        msg = fn()
        results.append((True, name, msg or ""))
        print(f"  {PASS} {name}" + (f" — {msg}" if msg else ""))
    except Exception as e:
        results.append((False, name, str(e)))
        print(f"  {FAIL} {name}: {e}")

# ── 1. File structure ─────────────────────────────────────────────────────────
print("\n── File structure ──────────────────────────────────────────────")

REQUIRED_FILES = [
    BASE / "app.py",
    BASE / "shared.py",
    BASE / "page_citation_repair.py",
    BASE / "page_library_tools.py",
    BASE / "page_bibliography_audit.py",
    BASE / "page_finalise.py",
    BASE / "page_figure_inventory.py",
]
# App blocks are indented code snippets (exec'd inside tab contexts) — check they exist but not syntax
BLOCK_FILES = [BLOCKS / f"app{i}.py" for i in range(1, 12)]

for path in REQUIRED_FILES:
    def _check(p=path):
        src = p.read_text()
        ast.parse(src)
        return f"{len(src.splitlines())} lines"
    test(f"Syntax: {path.relative_to(BASE)}", _check)

def check_blocks_exist():
    missing = [f.name for f in BLOCK_FILES if not f.exists()]
    if missing: raise AssertionError(f"Missing blocks: {missing}")
    return f"All {len(BLOCK_FILES)} app blocks present"
test("_app_blocks/ — all 11 files present", check_blocks_exist)

# ── 2. Routing ────────────────────────────────────────────────────────────────
print("\n── Routing & navigation ────────────────────────────────────────")

def check_routing():
    src = (BASE / "app.py").read_text()
    pages = ["citation_repair","library_tools","bibliography_audit",
             "finalise","figure_inventory"]
    missing = [p for p in pages if p not in src]
    if missing: raise AssertionError(f"Missing routes: {missing}")
    if "st.rerun()" not in src: raise AssertionError("st.rerun() missing")
    if 'session_state["page"]' not in src: raise AssertionError("page state missing")
    return f"{len(pages)} pages routed"
test("All pages routed in app.py", check_routing)

def check_tabs():
    checks = [
        ("page_citation_repair.py",
         ["Broken Citation Fixer","Document Merger","Citation Repair"]),
        ("page_library_tools.py",
         ["PubMed Search","Reference Comparator","RecNum Inspector"]),
        ("page_finalise.py",
         ["Citation Renumbering","Document Health Check","Batch Rename"]),
    ]
    for fn, expected in checks:
        src = (BASE / fn).read_text()
        missing = [t for t in expected if t not in src]
        if missing: raise AssertionError(f"{fn} missing tabs: {missing}")
    return "All tab labels present"
test("Tab labels correct in grouped pages", check_tabs)

def check_standalone_pages():
    # Standalone pages should not open with a top-level tab bar for navigation
    # (they may still use st.tabs() internally for results display — that's fine)
    for fn in ["page_bibliography_audit.py", "page_figure_inventory.py"]:
        page_src = (BASE / fn).read_text()
        lines = page_src.splitlines()
        # Check the first 15 lines don't have a navigation-level st.tabs call
        top_tabs = [l for l in lines[:15] if "st.tabs(" in l]
        if top_tabs:
            raise AssertionError(f"{fn} has top-level navigation tabs")
    return "Standalone pages have no top-level nav tabs"
test("Standalone pages have no nav-level tabs", check_standalone_pages)

# ── 3. Shared functions ───────────────────────────────────────────────────────
print("\n── Shared functions (shared.py) ────────────────────────────────")

def load_shared():
    src = (BASE / "shared.py").read_text()
    stubs = (
        "import streamlit as st\n"
        "class _ST:\n"
        "    session_state = type('S',(),{"
        "'__getattr__':lambda s,n:None,"
        "'__setattr__':lambda s,n,v:None,"
        "'__contains__':lambda s,n:False"
        "})()\n"
        "    def __getattr__(self,n): return lambda *a,**k: None\n"
        "st = _ST()\n"
    )
    ns = {}
    exec(stubs + src, ns)
    return ns

def check_core_functions():
    ns = load_shared()
    needed = ["fix_broken_fields", "remove_orphan_superscripts",
              "safe_merge_documents", "analyze_merge_damage",
              "analyze_docx_citations"]
    missing = [f for f in needed if f not in ns]
    if missing: raise AssertionError(f"Missing: {missing}")
    return f"{len(needed)} functions present"
test("Core repair functions in shared.py", check_core_functions)

def check_app_css():
    ns = load_shared()
    if "APP_CSS" not in ns: raise AssertionError("APP_CSS not exported")
    css = ns["APP_CSS"]
    for token in ["Libre Baskerville", "--accent:", "--bg:", "sidebar-logo"]:
        if token not in css: raise AssertionError(f"CSS missing: {token}")
    return f"APP_CSS present ({len(css):,} chars)"
test("APP_CSS exported and complete", check_app_css)

def check_pattern_a():
    ns = load_shared()
    data = ('<EndNote><Cite><Author>Smith</Author><Year>2020</Year>'
            '<RecNum>42</RecNum><record><rec-number>42</rec-number>'
            '<foreign-keys><key app="EN" db-id="testdb123">42</key></foreign-keys>'
            '<ref-type name="Journal Article">17</ref-type>'
            '<contributors><authors><author>Smith, J</author></authors></contributors>'
            '<titles><title>Test</title></titles>'
            '<dates><year>2020</year></dates></record></Cite></EndNote>')
    b64 = base64.b64encode(data.encode()).decode()
    xml = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           f'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
           f'<w:body><w:p>'
           f'<w:r><w:instrText xml:space="preserve"> ADDIN EN.CITE </w:instrText></w:r>'
           f'<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
           f'<w:r><w:t></w:t></w:r>'
           f'<w:r><w:fldChar w:fldCharType="end"/></w:r>'
           f'<w:fldData xml:space="preserve">{b64}</w:fldData>'
           f'</w:p></w:body></w:document>')
    fixed, n = ns["fix_broken_fields"](xml)
    if n == 0: raise AssertionError("No fields fixed")
    if "&lt;RecNum&gt;42&lt;/RecNum&gt;" not in fixed:
        raise AssertionError("RecNum 42 missing from output")
    return f"Fixed {n} field(s), RecNum 42 present"
test("fix_broken_fields — Pattern A repair", check_pattern_a)

def check_orphan_removal():
    ns = load_shared()
    xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
           '<w:body><w:p>'
           '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
           '<w:r><w:instrText> ADDIN EN.CITE &lt;EndNote&gt;&lt;Cite&gt;'
           '&lt;RecNum&gt;5&lt;/RecNum&gt;&lt;/Cite&gt;&lt;/EndNote&gt;</w:instrText></w:r>'
           '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
           '<w:r><w:t>5</w:t></w:r>'
           '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
           '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t>5, 6</w:t></w:r>'
           '</w:p></w:body></w:document>')
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        z.writestr("word/document.xml", xml)
        z.writestr("[Content_Types].xml",
            '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
            '</Types>')
        z.writestr("_rels/.rels",
            '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
            '</Relationships>')
        z.writestr("word/_rels/document.xml.rels",
            '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>')
    buf.seek(0)
    fixed_bytes, n = ns["remove_orphan_superscripts"](buf.read())
    if n == 0: raise AssertionError("No orphans removed")
    return f"Removed {n} orphan run(s)"
test("remove_orphan_superscripts — orphan removed", check_orphan_removal)

# ── 4. CSS consistency ────────────────────────────────────────────────────────
print("\n── CSS consistency ─────────────────────────────────────────────")

def check_css_classes():
    shared_src = (BASE / "shared.py").read_text()
    css_m = re.search(r"<style>([\s\S]+?)</style>", shared_src)
    if not css_m: raise AssertionError("No <style> block found")
    css = css_m.group(1)
    defined = set(re.findall(r"\.([\w-]+)\s*\{", css))
    all_used = set()
    for path in REQUIRED_FILES:
        src = path.read_text()
        for cls_str in re.findall(r'class=["\']([^"\']+)["\']', src):
            for c in cls_str.split():
                if "-" in c:
                    all_used.add(c)
    dynamic = {"active","done","waiting","good","bad","warn","error",
               "warning","ok","high","mid","low","accepted","skipped"}
    missing = all_used - defined - dynamic
    if missing: raise AssertionError(f"CSS missing: {missing}")
    return f"{len(defined)} defined, all {len(all_used)} custom classes covered"
test("All CSS classes defined in shared.py", check_css_classes)

def check_css_vars():
    css_m = re.search(r"<style>([\s\S]+?)</style>",
                      (BASE / "shared.py").read_text())
    css = css_m.group(1)
    declared = set(re.findall(r"--([\w-]+):", css))
    used      = set(re.findall(r"var\(--([\w-]+)\)", css))
    missing   = used - declared
    if missing: raise AssertionError(f"Vars used but not declared: {missing}")
    return f"{len(declared)} vars declared"
test("CSS variables all declared", check_css_vars)

# ── 5. UI quality checks ──────────────────────────────────────────────────────
print("\n── UI quality ──────────────────────────────────────────────────")

def check_no_step_concat():
    for path in REQUIRED_FILES:
        bad = re.findall(r"STEP \d[A-Z]", path.read_text())
        if bad: raise AssertionError(f"{path.name}: {bad}")
    return "No STEP+Title concatenation"
test("No STEP+Title concatenation", check_no_step_concat)

def check_button_colors():
    css_m = re.search(r"<style>([\s\S]+?)</style>",
                      (BASE / "shared.py").read_text())
    css = css_m.group(1)
    bad = ["background: #000", "background: black", "background: #111"]
    for bc in bad:
        if bc in css: raise AssertionError(f"Dark button: {bc}")
    if "#8b1a1a" not in css and "#3d5a6e" not in css:
        raise AssertionError("No button color found")
    return "Button colors readable"
test("Button colors not dark-on-dark", check_button_colors)

def check_app_labels():
    # Every page file should have an app-label div
    for fn in ["page_citation_repair.py","page_library_tools.py",
               "page_bibliography_audit.py","page_finalise.py",
               "page_figure_inventory.py"]:
        src = (BASE / fn).read_text()
        if "app-label" not in src:
            raise AssertionError(f"{fn} missing app-label div")
    return "All page files have app-label"
test("App-label badges on all pages", check_app_labels)

# ── Summary ───────────────────────────────────────────────────────────────────
print()
passed = sum(1 for r in results if r[0])
failed = sum(1 for r in results if not r[0])
total  = len(REQUIRED_FILES)
print(f"── Results: {passed} passed, {failed} failed "
      f"({'✓' if not failed else '✗'})")
all_files = REQUIRED_FILES + BLOCK_FILES
print(f"   {len(all_files)} files checked ({sum(len(f.read_text().splitlines()) for f in all_files if f.exists())} total lines)\n")
if failed:
    print("Failed:")
    for ok, name, msg in results:
        if not ok: print(f"  ✗ {name}: {msg}")
    sys.exit(1)
else:
    print("   All tests passed — safe to upload to GitHub.\n")
