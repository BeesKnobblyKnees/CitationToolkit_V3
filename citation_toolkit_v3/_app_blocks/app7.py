    st.markdown('<div class="app-label">A practical tool &nbsp;07</div>', unsafe_allow_html=True)
    st.markdown("## Citation Renumbering")
    st.markdown('''<div class="instruction-box">
    <b>When to use:</b> After editing, citation superscript numbers and the bibliography
    are out of order. This tool renumbers them consistently using your chosen method.<br><br>
    <b>Before using:</b> Make sure EndNote has already formatted the bibliography so
    citations appear as plain superscript numbers in the text — not live field codes.
    </div>''', unsafe_allow_html=True)

    ren_file = st.file_uploader(
        "Word document (.docx) with formatted citations",
        type=["docx"], key="ren_doc"
    )

    method = st.radio(
        "Numbering method",
        [
            "Alphabetical — A=1, B=2... (sort bibliography by author last name)",
            "Order of appearance — first cited in text = 1, second = 2...",
        ]
    )

    if ren_file:
        if st.button("Renumber citations", type="primary"):
            with st.spinner("Scanning and renumbering..."):
                raw_bytes = ren_file.read()
                if "Alphabetical" in method:
                    fixed_bytes, mapping = renumber_citations_alpha(raw_bytes)
                    method_label = "alphabetically (A-Z by author)"
                else:
                    fixed_bytes, mapping = renumber_citations_appearance(raw_bytes)
                    method_label = "by order of appearance"

            if not mapping:
                st.warning(
                    "No superscript citation numbers or numbered bibliography found. "
                    "Make sure EndNote has formatted the bibliography first."
                )
            else:
                changed = {k: v for k, v in mapping.items() if k != v}
                st.success(
                    f"Done — renumbered {method_label}. "
                    f"{len(mapping)} unique citations, {len(changed)} numbers changed."
                )
                col1, col2, col3 = st.columns(3)
                col1.metric("Unique citations",  len(mapping))
                col2.metric("Numbers changed",   len(changed))
                col3.metric("Already in order",  len(mapping) - len(changed))

                st.download_button(
                    "⬇ Download renumbered document",
                    data=fixed_bytes,
                    file_name=Path(ren_file.name).stem + "_renumbered.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    type="primary"
                )

                if changed:
                    with st.expander(f"Renumbering map ({len(changed)} changes)"):
                        col1, col2 = st.columns(2)
                        col1.markdown("**Old #**")
                        col2.markdown("**New #**")
                        for old in sorted(mapping):
                            new = mapping[old]
                            if old != new:
                                col1.markdown(str(old))
                                col2.markdown(str(new))

# ─────────────────────────────────────────────────────────────────────────────
# APP 6 UI — FIGURE INVENTORY
