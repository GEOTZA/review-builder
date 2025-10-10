# μετά το file_uploader:
if run:
    import time

    if not xls:
        st.error("Ανέβασε Excel πρώτα.")
        st.stop()
    if not tpl_bex or not tpl_nonbex:
        st.error("Ανέβασε και τα δύο templates.")
        st.stop()

    st.info(f"📄 Excel size: {len(xls.getbuffer())/1024:.1f} KB | BEX tpl: {tpl_bex.size/1024:.1f} KB | Non-BEX tpl: {tpl_nonbex.size/1024:.1f} KB")

    # Δείξε spinner για την ανάγνωση Excel
    with st.spinner("Ανάγνωση Excel..."):
        try:
            df = pd.read_excel(xls, sheet_name=sheet_name, engine="openpyxl")
        except Exception as e:
            st.error(f"Δεν άνοιξε το Excel (sheet='{sheet_name}'): {e}")
            st.stop()

    st.success(f"OK: βρέθηκαν {len(df)} γραμμές και {len(df.columns)} στήλες.")
    st.dataframe(df.head(5))

    cols = list(df.columns)

    # ... (κρατάς το auto-mapping όπως το έχουμε) ...

    # Templates σε μνήμη
    tpl_bex_bytes = tpl_bex.read()
    tpl_nonbex_bytes = tpl_nonbex.read()

    out_zip = io.BytesIO()
    z = zipfile.ZipFile(out_zip, "w", zipfile.ZIP_DEFLATED)
    built = 0

    pbar = st.progress(0, text="Δημιουργία εγγράφων...")
    total = max(1, len(df))

    def cell(row, col):
        if not col: return ""
        v = row[col]
        if pd.isna(v): return ""
        return v

    for i, (_, row) in enumerate(df.iterrows(), start=1):
        try:
            store = str(cell(row, col_store)).strip()
            if not store:
                pbar.progress(min(i/total, 1.0), text=f"Παράλειψη γραμμής {i} (κενό store)")
                continue
            store_up = store.upper()

            if bex_mode == "Λίστα (comma-separated)":
                is_bex = store_up in bex_list
            else:
                bex_val = str(cell(row, col_bex)).strip().lower()
                is_bex = bex_val in ("yes","y","1","true","ναι")

            mapping = {
                "title": f"Review September 2025 — Plan October 2025 — {store_up}",
                "store": store_up,
                "mobile_actual":  cell(row, col_mob_act),
                "mobile_target":  cell(row, col_mob_tgt),
                "fixed_actual":   cell(row, col_fix_act),
                "fixed_target":   cell(row, col_fix_tgt),
                "pending_mobile": cell(row, col_pend_mob),
                "pending_fixed":  cell(row, col_pend_fix),
                "plan_vs_target": cell(row, col_plan_vs),
            }

            tpl_bytes = tpl_bex_bytes if is_bex else tpl_nonbex_bytes
            doc = Document(io.BytesIO(tpl_bytes))
            set_default_font(doc, "Aptos")
            replace_placeholders(doc, mapping)

            out_name = f"{store_up}_ReviewSep_PlanOct.docx"
            buf = io.BytesIO()
            doc.save(buf)
            z.writestr(out_name, buf.getvalue())
            built += 1
            pbar.progress(min(i/total, 1.0), text=f"Φτιάχνω: {out_name} ({i}/{total})")
        except Exception as e:
            st.warning(f"⚠️ Σφάλμα στη γραμμή {i}: {e}")
            pbar.progress(min(i/total, 1.0), text=f"Συνεχίζω… ({i}/{total})")

    z.close()
    if built == 0:
        st.error("Δεν δημιουργήθηκε αρχείο. Έλεγξε αν αναγνωρίστηκε η στήλη STORE και τα templates.")
    else:
        st.success(f"Έτοιμα {built} αρχεία.")
        st.download_button("⬇️ Κατέβασε ZIP", data=out_zip.getvalue(), file_name="reviews_from_excel.zip")
