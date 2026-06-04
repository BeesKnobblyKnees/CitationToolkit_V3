#!/usr/bin/env python3
"""
fix_expander_css_v2.py
Replaces the expander block added by fix_expander_css.py with a corrected one.

Root cause found via Inspect Element: the summary was 970x1274px (should be ~45px).
The `text-indent: -9999px` trick was inherited into the inner label <span>, collided
with the flex layout, and squished the content into a tall column where the arrow
ligature leaked as "_arr". This version removes the text-indent trick entirely,
neutralises the Safari ::before cover, and lays the summary > span out normally.

Run from inside your toolkit folder:
    python3 fix_expander_css_v2.py
"""

from pathlib import Path
import shutil
import sys

SHARED = Path("shared.py")
if not SHARED.exists():
    print("❌ shared.py not found. Run this from inside citation_toolkit_v3/")
    sys.exit(1)

content = SHARED.read_text()
shutil.copy(SHARED, "shared.py.bak2")
print("✓ Backed up shared.py -> shared.py.bak2")

START = "/* ===================================================================\n   EXPANDERS — single source of truth"
END = "/* ===================== END EXPANDERS ===================== */"

NEW_BLOCK = """/* ===================================================================
   EXPANDERS — single source of truth (v2). Edit only here.
   =================================================================== */
div[data-testid="stExpander"] {
  overflow: hidden !important;
  border-radius: var(--radius) !important;
  border: 1px solid var(--border) !important;
  background: var(--surface) !important;
  margin-bottom: 0.6rem !important;
}

/* Header row — NO text-indent trick (that caused the 1274px-tall button). */
div[data-testid="stExpander"] details > summary {
  display: flex !important;
  align-items: center !important;
  gap: 8px !important;
  padding: 0.75rem 1rem !important;
  cursor: pointer !important;
  background: var(--surface) !important;
  list-style: none !important;
  user-select: none !important;
  text-indent: 0 !important;
  min-height: 0 !important;
  height: auto !important;
  position: static !important;   /* undo old Safari position:relative */
  padding-left: 1rem !important; /* undo old Safari padding-left:32px */
}
div[data-testid="stExpander"] details > summary:hover {
  background: var(--bg2) !important;
}

/* Remove native disclosure markers AND the old Safari cover box. */
div[data-testid="stExpander"] details > summary::-webkit-details-marker { display: none !important; }
div[data-testid="stExpander"] details > summary::marker { display: none !important; content: "" !important; }
div[data-testid="stExpander"] details > summary::before { content: none !important; display: none !important; }

/* The single label span that holds the arrow + text — lay it out normally. */
div[data-testid="stExpander"] details > summary > span {
  display: flex !important;
  align-items: center !important;
  gap: 8px !important;
  text-indent: 0 !important;
  white-space: normal !important;
  overflow: visible !important;
  width: auto !important;
  color: var(--ink) !important;
  font-family: 'Source Sans 3', sans-serif !important;
  font-size: 0.92rem !important;
  font-weight: 500 !important;
}

/* Arrow handling: if it's a real <svg>, show it; if it's an icon-font glyph
   (the thing leaking "_arr"), zero its font so the stray text can't render. */
div[data-testid="stExpander"] details > summary svg {
  font-size: 1rem !important;
  width: 16px !important;
  height: 16px !important;
  flex-shrink: 0 !important;
  color: var(--ink-dim) !important;
  fill: var(--ink-dim) !important;
}
div[data-testid="stExpander"] details > summary > span > span:not(:has(svg)),
div[data-testid="stExpander"] details > summary [class*="material-symbols"],
div[data-testid="stExpander"] details > summary [data-testid="stIconMaterial"] {
  font-size: 0 !important;
  line-height: 0 !important;
  color: transparent !important;
}

/* Body text inside the expander. */
div[data-testid="stExpander"] div[data-testid="stMarkdownContainer"] p,
div[data-testid="stExpander"] div[data-testid="stMarkdownContainer"] * {
  color: var(--ink-mid) !important;
  font-family: 'Source Sans 3', sans-serif !important;
}
/* ===================== END EXPANDERS ===================== */"""

s = content.find(START)
e = content.find(END)
if s == -1 or e == -1:
    print("❌ Could not find the v1 expander block markers.")
    print("   Did fix_expander_css.py run? No changes written; backup at shared.py.bak2")
    sys.exit(1)

e_full = e + len(END)
content = content[:s] + NEW_BLOCK + content[e_full:]
SHARED.write_text(content)
print("✓ Replaced expander block with v2")
print("\n✅ Done. Refresh localhost and check Citation Repair.")
print("   The giant button should now be normal height.")
print("   Undo if needed:  mv shared.py.bak2 shared.py")
