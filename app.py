# app.py
# Streamlit: Excel/CSV → BEX & Non-BEX Review/Plan Generator
# Author: GEOTZA + Κλέρη ❤️

import io, re, zipfile
import pandas as pd
import streamlit as st
from typing import Any, Dict
from docx import Document
from docx.oxml.ns import qn

# ───────────────────────────── CONFIG ─────────────────────────────
st.set_page_config(page_title="Excel → Review/Plan Generator", layout="wide")
st.title("📊 Excel/CSV → 📄 Review/Plan Generator (BEX & Non-BEX)")

# ───────────────────────────── HELPERS ─────────────────────────────
def set_default_font(doc: Document, font_name="Aptos"):
    for style in doc.styles:
        if hasattr(style, "font"):
            try:
                style.font.name = font_name
                style._element.rPr.rFonts.set(qn("w:eastAsia"), font_name)
                style._element.rPr.rFonts.set(qn("w:cs"), font_name)
            except Exception:
                pass

def replace_placeholders(doc: Document, mapping: Dict[str, Any]):
    pattern = re.compile(r"\[\[([A-Za-z0-9_]+)\]\]")
    def subfun(s: str) -> str:
        return pattern.sub(lambda m: str(mapping.get(m.group(1), "")), s)
    for p in doc.paragraphs:
        for r in p.runs:
            r.text = subfun(r.text)
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        r.text = subfun(r.text)

def excel_letter_to_colname(df, letter: str):
    """Μετατρέπει γράμμα Excel σε όνομα στήλης DataFrame."""
    try:
        idx = pd.ExcelFile._col_letter_to_index(letter)
        return df.columns[idx]
    except Exception:
        return None

def cell(row, col):
    if not col:
        return ""
    v = row[col]
    return "" if pd.isna(v) else v

def read_data(xls, sheet_name):
    try:
        if xls.name.lower().endswith(".csv"):
            st.write("📑 Sheets:", ["CSV Data"])
            return pd.read_csv(xls)
        xfile = pd.ExcelFile(xls, engine="openpyxl")
        st.write("📑 Sheets:", xfile.sheet_names)
        if sheet_name not in xfile.sheet_names:
            st.error(f"Το sheet '{sheet_name}' δεν βρέθηκε.")
            return None
        return pd.read_excel(xfile, sheet_name=sheet_name, engine="openpyxl")
    except Exception as e:
        st.error(f"Δεν άνοιξε το αρχείο: {e}")
        return None

# ───────────────────────────── SIDEBAR ─────────────────────────────
debug_mode = st.sidebar.toggle("🛠 Debug mode", value=True)
test_mode  = st.sidebar.toggle("🧪 Test mode (limit rows=50)", value=True)

st.sidebar.header("⚙️ BEX Settings")
bex_stores_fixed = {"DRZ01", "FKM01", "ESC01", "LND01", "PKK01"}
st.sidebar.info("Τα BEX stores: DRZ01, FKM01, ESC01, LND01, PKK01")

st.sidebar.subheader("📄 Templates (.docx)")
tpl_bex = st.sidebar.file_uploader("BEX template", type=["docx"])
tpl_nonbex = st.sidebar.file_uploader("Non-BEX template", type=["docx"])
st.sidebar.caption("Placeholders: [[title]], [[store]], [[plan_vs_target]], [[mobile_actual]], [[mobile_target]], [[fixed_actual]], [[fixed_target]], [[pending_mobile]], [[pending_fixed]]")

# 🧩 Manual Mapping
st.sidebar.subheader("📎 Manual mapping (Excel letters)")
letter_plan_vs    = st.sidebar.text_input("plan vs target", "A")
letter_mobile_plan = st.sidebar.text_input("mobile plan (optional)", "B")
letter_bex_col    = st.sidebar.text_input("BEX (YES/NO) column", "J")
letter_mobile_act = st.sidebar.text_input("mobile actual", "N")
letter_mobile_tgt = st.sidebar.text_input("mobile target", "O")
letter_fixed_tgt  = st.sidebar.text_input("fixed target", "P")
letter_fixed_act  = st.sidebar.text_input("total fixed actual", "Q")
letter_voice_vs   = st.sidebar.text_input("voice vs target", "R")
letter_fixed_vs   = st.sidebar.text_input("fixed vs target", "S")
letter_llu        = st.sidebar.text_input("llu actual", "T")
letter_nga        = st.sidebar.text_input("nga actual", "U")
letter_ftth       = st.sidebar.text_input("ftth actual", "V")
letter_eon        = st.sidebar.text_input("eon tv actual", "X")
letter_fwa        = st.sidebar.text_input("fwa actual", "Y")
letter_mob_upg    = st.sidebar.text_input("mobile upgrades", "AA")
letter_fix_upg    = st.sidebar.text_input("fixed upgrades", "AB")
letter_pend_mob   = st.sidebar.text_input("total pending mobile", "AF")
letter_pend_fix   = st.sidebar.text_input("total pending fixed", "AH")

letters_map = {
    "plan_vs_target": letter_plan_vs,
    "mobile_plan": letter_mobile_plan,
    "mobile_actual": letter_mobile_act,
    "mobile_target": letter_mobile_tgt,
    "fixed_target": letter_fixed_tgt,
    "fixed_actual": letter_fixed_act,
    "voice_vs_target": letter_voice_vs,
    "fixed_vs_target": letter_fixed_vs,
    "llu_actual": letter_llu,
    "nga_actual": letter_nga,
    "ftth_actual": letter_ftth,
    "eon_tv_actual": letter_eon,
    "fwa_actual": letter_fwa,
    "mobile_upgrades": letter_mob_upg,
    "fixed_upgrades": letter_fix_upg,
    "pending_mobile": letter_pend_mob,
    "pending_fixed": letter_pend_fix,
}

# ───────────────────────────── MAIN ─────────────────────────────
st.markdown("### 1) Ανέβασε Excel/CSV")
xls = st.file_uploader("Excel/CSV", type=["xlsx", "csv"])
sheet_name = st.text_input("Όνομα φύλλου (Sheet - μόνο για Excel)", value="Sheet1")

run = st.button("🔧 Generate")

if run:
    if not xls:
        st.error("Ανέβασε αρχείο Excel ή CSV πρώτα.")
        st.stop()
    if not tpl_bex or not tpl_nonbex:
        st.error("Ανέβασε και τα δύο templates (.docx).")
        st.stop()

    df = read_data(xls, sheet_name)
    if df is None or df.empty:
        st.error("Δεν βρέθηκαν δεδομένα στο αρχείο.")
        st.stop()

    st.success(f"OK: {len(df)} γραμμές, {len(df.columns)} στήλες.")
    if debug_mode:
        st.dataframe(df.head(10))

    # Live preview των headers
    headers_preview = {}
    for k, L in letters_map.items():
        hdr = excel_letter_to_colname(df, L) if L and L.strip() else None
        headers_preview[k] = hdr
    with st.expander("🧭 Letters → Headers (live)"):
        st.json(headers_preview)

    # Safety check
    conflicts = [k for k, v in headers_preview.items() if v == "Dealer_Code"]
    if conflicts:
        st.warning(f"⚠️ Πεδία που πέφτουν σε λάθος στήλη: {conflicts}")

    # Templates
    tpl_bex_bytes = tpl_bex.read()
    tpl_nonbex_bytes = tpl_nonbex.read()
    out_zip = io.BytesIO()
    zf = zipfile.ZipFile(out_zip, "w", zipfile.ZIP_DEFLATED)
    built = 0
    total = len(df) if not test_mode else min(50, len(df))
    pbar = st.progress(0, text="Δημιουργία εγγράφων…")

    for i, (_, row) in enumerate(df.iterrows(), start=1):
        if test_mode and i > total:
            break
        try:
            store = str(row["Dealer_Code"]).strip()
            if not store:
                continue
            is_bex = store.upper() in bex_stores_fixed
            mapping = {
                "title": f"Review September 2025 — Plan October 2025 — {store}",
                "store": store,
            }
            for k, L in letters_map.items():
                hdr = excel_letter_to_colname(df, L)
                mapping[k] = cell(row, hdr)
            doc = Document(io.BytesIO(tpl_bex_bytes if is_bex else tpl_nonbex_bytes))
            set_default_font(doc, "Aptos")
            replace_placeholders(doc, mapping)
            buf = io.BytesIO()
            doc.save(buf)
            zf.writestr(f"{store}_ReviewSep_PlanOct.docx", buf.getvalue())
            built += 1
            pbar.progress(min(i / (total or 1), 1.0), text=f"Φτιάχνω: {store} ({i}/{total})")
        except Exception as e:
            st.warning(f"⚠️ Σφάλμα στη γραμμή {i}: {e}")
            if debug_mode:
                st.exception(e)
    zf.close()
    pbar.empty()

    if built == 0:
        st.error("Δεν δημιουργήθηκε αρχείο.")
    else:
        st.success(f"✅ Έτοιμα {built} αρχεία.")
        st.download_button("⬇️ Κατέβασε ZIP", data=out_zip.getvalue(), file_name="reviews_from_excel.zip")