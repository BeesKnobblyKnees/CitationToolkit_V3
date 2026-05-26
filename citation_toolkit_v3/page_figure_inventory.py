"""
Standalone page.
"""
import streamlit as st, zipfile, re, base64, io, html
from pathlib import Path
from shared import *

st.markdown(APP_CSS, unsafe_allow_html=True)
st.divider()

st.markdown("## Figure & Table Inventory")
st.markdown('<div class="instruction-box">Scans your document for all figures, tables, boxes, plates, and videos. Optionally cross-references against your Excel naming sheet to flag mismatches and missing items.</div>',unsafe_allow_html=True)
col1,col2=st.columns(2)
with col1: fig_doc=st.file_uploader("Word document (.docx)",type=["docx"],key="fig_doc")
with col2: fig_excel=st.file_uploader("Excel naming sheet (optional)",type=["xlsx","xls"],key="fig_excel",help="Columns: Number | Name/Caption")
if fig_doc:
    if st.button("Scan document",type="primary"):
        with st.spinner("Scanning..."):
            fig_bytes=fig_doc.read(); items=scan_figures(fig_bytes)
        if not items:
            st.warning("No captioned items found. Captions must start with Figure/Table/Box/Plate/Video followed by a number.")
        else:
            type_counts={}
            for item in items: type_counts[item['type']]=type_counts.get(item['type'],0)+1
            cols=st.columns(min(len(type_counts),6))
            for i,(t,c) in enumerate(sorted(type_counts.items())): cols[i%len(cols)].metric(t+"s",c)
            st.markdown(f"**{len(items)} total items**")
            if fig_excel:
                excel_bytes=fig_excel.read(); results=cross_ref_excel(items,excel_bytes)
                matched=[r for r in results if r[1]=='match']
                mismatch=[r for r in results if r[1]=='mismatch']
                not_found=[r for r in results if r[1]=='not_in_excel']
                col1,col2,col3,col4=st.columns(4)
                col1.metric("Matched",len(matched)); col2.metric("Mismatch",len(mismatch))
                col3.metric("Not in Excel",len(not_found))
                tabs=st.tabs([f"Mismatches ({len(mismatch)})",f"Not in Excel ({len(not_found)})",
                              f"Matched ({len(matched)})",f"All ({len(results)})"])
                def show_res(res_list):
                    for item,status,expected in res_list:
                        color='ok' if status=='match' else 'missing' if status=='not_in_excel' else 'warn'
                        exp=""
                        if expected and status=='mismatch': exp=f'<br><span style="font-size:0.75rem;color:#ffb300">Expected: {expected}</span>'
                        st.markdown(f'<div class="ref-item {color}"><b>{item["type"]} {item["number"]}</b>{"🖼" if item["has_image"] else ""} — Para {item["para_idx"]+1}<br><span style="font-size:0.82rem">{item["caption"][:180]}</span>{exp}</div>',unsafe_allow_html=True)
                with tabs[0]: show_res(mismatch) if mismatch else st.success("No mismatches.")
                with tabs[1]: show_res(not_found) if not_found else st.success("All found in Excel.")
                with tabs[2]: show_res(matched) if matched else st.info("No confirmed matches.")
                with tabs[3]: show_res(results)
            else:
                type_filter=st.selectbox("Filter by type",["All"]+sorted(type_counts.keys()))
                for item in items:
                    if type_filter!="All" and item['type']!=type_filter: continue
                    st.markdown(f'<div class="match-card"><span class="match-marker">{item["type"]} {item["number"]}</span>{"🖼" if item["has_image"] else ""}<span style="font-size:0.72rem;color:#3a4a5a"> Para {item["para_idx"]+1}</span><div class="match-sentence">{item["caption"][:200]}</div></div>',unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# APP 7 UI — PUBMED SEARCH
