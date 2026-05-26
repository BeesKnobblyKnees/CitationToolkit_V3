    st.markdown('<div class="app-label">A practical tool &nbsp;03</div>', unsafe_allow_html=True)
    st.markdown("## Document Merger")

    merge_mode = st.radio(
        "What do you need to do?",
        [
            "Restore lost citations — merge citation-intact OLD doc with text-edited NEW doc",
            "Repair already-merged document — fix broken field codes after a merge",
        ]
    )
    st.divider()

    if "Restore lost citations" in merge_mode:
        st.markdown('''<div class="instruction-box">
        <b>Use this when:</b> A merge broke or removed inline citations from the new document,
        but you still have the older version with all citations intact.<br><br>
        The app matches paragraphs between the two documents by text content and copies
        citation field codes from the old document into the new one wherever they are missing.
        All text edits in the new document are preserved.
        </div>''', unsafe_allow_html=True)

        col1, col2 = st.columns(2)
        with col1:
            st.markdown("**New document** — has text edits but missing/broken citations")
            new_file = st.file_uploader("New .docx", type=["docx"], key="merge_new")
        with col2:
            st.markdown("**Old document** — has all citations intact")
            old_file = st.file_uploader("Old .docx", type=["docx"], key="merge_old")

        if new_file and old_file:
            new_bytes_m = new_file.read()
            old_bytes_m = old_file.read()
            new_fname   = new_file.name

            new_analysis = analyze_merge_damage(new_bytes_m)
            col1, col2, col3 = st.columns(3)
            col1.metric("Bibliography entries",  new_analysis.get("bib_count", 0))
            col2.metric("Cited in text (new)",   new_analysis.get("cited_count", 0))
            col3.metric("Citations lost",         new_analysis.get("lost_in_merge", 0),
                        delta=f'-{new_analysis.get("lost_in_merge",0)}' if new_analysis.get("lost_in_merge",0) else None,
                        delta_color="inverse")

            if st.button("Restore citations from old document", type="primary"):
                with st.spinner("Matching paragraphs and restoring citations..."):
                    try:
                        merged_bytes, rpt = safe_merge_documents(new_bytes_m, old_bytes_m)
                        st.success(
                            f"Done. Matched {rpt['matched']} paragraphs. "
                            f"Restored citations in {rpt['citations_restored']} paragraph(s). "
                            f"{rpt['unmatched']} paragraph(s) could not be matched to the old document."
                        )
                        col1,col2,col3,col4 = st.columns(4)
                        col1.metric("Matched",           rpt["matched"])
                        col2.metric("Citations restored",rpt["citations_restored"])
                        col3.metric("Already had cites", rpt["already_had_cites"])
                        col4.metric("Unmatched",         rpt["unmatched"])
                        if rpt["unmatched"] > 0:
                            st.info(
                                f"{rpt['unmatched']} paragraph(s) are new content not in the old document "
                                f"— these could not have citations restored automatically."
                            )
                        st.download_button(
                            "⬇ Download merged document",
                            data=merged_bytes,
                            file_name=Path(new_fname).stem + "_citations_restored.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            type="primary"
                        )
                        st.markdown('''<div class="instruction-box">
                        <b>After downloading:</b><br>
                        1. Open in Word with your EndNote library connected<br>
                        2. EndNote tab → Update Citations and Bibliography<br>
                        3. If any citations still unlinked → use App 2 Remap tool<br>
                        4. Check any unmatched paragraphs manually
                        </div>''', unsafe_allow_html=True)
                        if rpt["details"]:
                            with st.expander(f"Paragraphs with restored citations ({len(rpt['details'])})"):
                                for d in rpt["details"]:
                                    st.markdown(
                                        f'<div class="ref-item ok">Para {d["para_idx"]+1} — {d["cites_added"]} citation run(s) restored<br><span style="font-size:0.8rem;font-style:italic">{d["text_preview"]}</span></div>',
                                        unsafe_allow_html=True
                                    )
                    except Exception as e:
                        st.error(f"Merge failed: {e}")
                        st.exception(e)

    else:
        st.markdown('''<div class="instruction-box">
        <b>When to use:</b> You merged two Word documents and EndNote no longer
        recognizes the citations. This tool accepts tracked changes safely
        (rescuing any citations inside deleted text) and repairs broken citation field codes.<br><br>
        <b>Best practice for future merges:</b> Before using Word's Compare, go to
        EndNote tab → Convert Citations → Convert to Unformatted Citations.
        This turns field codes into plain text like {Hall, 1997 #18} which survives
        merging perfectly. Then after accepting changes, use Update Citations and Bibliography.
        </div>''', unsafe_allow_html=True)

        col1, col2 = st.columns(2)
        with col1:
            st.markdown("**Post-merge document** — the merged file with broken citations")
            merged_file = st.file_uploader("Merged .docx", type=["docx"], key="merge_merged")
        with col2:
            st.markdown("**Original document** (optional) — used to detect lost citations")
            orig_file = st.file_uploader("Original .docx", type=["docx"], key="merge_orig")

        if merged_file:
            if st.button("Analyze & repair", type="primary"):
                with st.spinner("Analyzing citation damage..."):
                    merged_bytes = merged_file.read()
                    orig_bytes   = orig_file.read() if orig_file else None
                    analysis     = analyze_merge_damage(merged_bytes)

                st.markdown("### Damage report")
                bib_count   = analysis.get("bib_count", 0)
                cited_count = analysis.get("cited_count", 0)
                lost_count  = analysis.get("lost_in_merge", 0)
                field_count = analysis["total_en"]
                sees        = analysis.get("endnote_sees", analysis["with_data"])
                misses      = analysis.get("endnote_misses", analysis["empty_cite"])

                col1,col2,col3,col4 = st.columns(4)
                col1.metric("Bibliography entries",    bib_count,
                            help="Number of references in the numbered reference list")
                col2.metric("EndNote currently sees",  sees,
                            help="Unique references EndNote can read from working field codes")
                col3.metric("Hidden in broken fields", misses,
                            delta=f"-{misses}" if misses else None,
                            delta_color="inverse" if misses else "off",
                            help="References locked in broken field codes — EndNote cannot count these")
                col4.metric("Lost during merge",       lost_count,
                            delta=f"-{lost_count}" if lost_count else None,
                            delta_color="inverse" if lost_count else "off",
                            help="References whose inline citations were removed by the merge")

                if misses > 0:
                    st.error(
                        f"⚠ **EndNote is undercounting your references.** "
                        f"It currently sees **{sees}** unique references but **{misses}** more "
                        f"are locked in broken citation field codes and invisible to EndNote — "
                        f"causing your bibliography to show fewer entries than expected. "
                        f"This is fixed automatically below."
                    )
                if lost_count > 0:
                    st.warning(
                        f"⚠ {lost_count} citation(s) were lost during the merge. "
                        f"Use **Restore lost citations** mode with the original document to recover them."
                    )
                if misses == 0 and lost_count == 0 and bib_count > 0:
                    st.success(f"✓ All {bib_count} bibliography references are intact and visible to EndNote.")
                if not analysis["balanced"]:
                    st.warning(
                        f"⚠ Unbalanced field markers "
                        f"(begin:{analysis['begins']}/sep:{analysis['separates']}/end:{analysis['ends']}) "
                        f"— some citation fields were split during merge."
                    )
                elif field_count > 0:
                    st.success("✓ Citation field boundaries are intact.")
                if len(analysis["db_ids"]) > 1:
                    st.info(f"Multiple library fingerprints found ({len(analysis['db_ids'])}) — citations from different libraries.")

                with st.spinner("Repairing..."):
                    fixed_bytes, rpt = repair_post_merge_citations(merged_bytes, orig_bytes)
                st.markdown("### Results")
                col1,col2,col3 = st.columns(3)
                col1.metric("Citations before", rpt["citations_before"])
                col2.metric("Citations after",  rpt["citations_after"])
                col3.metric("Steps applied",    len(rpt["steps"]))
                for step in rpt["steps"]:
                    if step == "track_changes_accepted":
                        st.markdown("- ✓ Tracked changes accepted safely")
                    elif "restored" in step:
                        n = step.split("_")[1]
                        st.markdown(f"- ✓ {n} broken field(s) restored from backup data")
                    elif "lost" in step:
                        n = step.split("_")[0]
                        st.markdown(f"- ⚠ {n} citation(s) lost in merge — use Restore mode to recover")
                if rpt.get("lost_rec_nums"):
                    st.warning(f"⚠ {len(rpt['lost_rec_nums'])} citation(s) from original not found after merge.")
                st.download_button(
                    "⬇ Download repaired document",
                    data=fixed_bytes,
                    file_name=Path(merged_file.name).stem + "_repaired.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    type="primary"
                )
                st.markdown('''<div class="instruction-box">
                <b>After downloading:</b><br>
                1. Open in Word with your EndNote library connected<br>
                2. EndNote tab → Update Citations and Bibliography<br>
                3. Still issues? → Use App 2 Remap tool
                </div>''', unsafe_allow_html=True)
