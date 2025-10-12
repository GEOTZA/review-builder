f# app.py
# Streamlit: Excel/CSV -> (BEX / Non-BEX) Review-Plan .docx (ZIP)
# Author: GEOTZA + Nova helper

import io
import re
import zipfile
from typing import Any, Dict

import pandas as pd
import streamlit as st
from docx import Document
from docx.oxml.ns import qn

# ───────────────────────────── UI CONFIG ─────────────────────────────
st.set_page_config(page_title="Excel → Review/Plan Generator", layout="wide")
st.title("📊 Excel/CSV → 📄 Review/Plan Generator (BEX & Non-BEX)")

# ───────────────────────────── HELPERS ─────────────────────────────
def set_default_font(doc: Document, font_name: str = "Aptos") -> None:
    """Ορίζει προεπιλεγμένη γραμματοσειρά σε όλα τα styles (και eastAsia)."""
    for style in doc.styles:
        if hasattr(style, "font"):
            try:
                style.font.name = font_name
                style._element.rPr.rFonts.set(qn("w:eastAsia"), font_name)
                style._element.rPr.rFonts.set(qn("w:cs"), font_name)
            except Exception:
                pass

def replace_placeholders(doc: Document, mapping: Dict[str, Any]) -> None:
    """Αντικαθιστά [[placeholders]] σε paragraphs & tables."""
    pattern = re.compile(r"\[\[([A-Za-z0-9_]+)\]\]")

    def subfun(s: str) -> str:
        key_to_val = lambda m: "" if mapping.get(m.group(1)) is None else str(mapping.get(m.group(1), ""))
        return pattern.sub(key_to_val, s)

    for p in doc.paragraphs:
        for r in p.runs:
            r.text = subfun(r.text)
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        r.text = subfun(r.text)

def normkey(x: str) -> str:
    """lower + αφαίρεση κενών/-,_,. για robust ταύτιση headers."""
    return re.sub(r"[\s\-_\.]+", "", str(x).strip().lower())

def pick(columns, *aliases) -> str:
    """Βρες στήλη με βάση aliases (πρώτα exact normalized, μετά regex contains)."""
    nmap = {normkey(c): c for c in columns}
    for a in aliases:
        if normkey(a) in nmap:
            return nmap[normkey(a)]
    for a in aliases:
        pat = re.compile(a, re.IGNORECASE)
        for c in columns:
            if re.search(pat, str(c)):
                return c
    return ""

def cell(row: pd.Series, col: str):
    if not col:
        return ""
    v = row[col]
    return "" if pd.isna(v) else v

def read_data(xls, sheet_name: str) -> pd.DataFrame | None:
    """Δέχεται .xlsx ή .csv (auto-detect από το όνομα). Επιστρέφει DataFrame ή None."""
    try:
        fname = getattr(xls, "name", "")
        if fname.lower().endswith(".csv"):
            st.write("📑 Sheets:", ["CSV Data"])
            return pd.read_csv(xls)
        # default: xlsx
        xfile = pd.ExcelFile(xls, engine="openpyxl")
        st.write("📑 Sheets:", xfile.sheet_names)
        if sheet_name not in xfile.sheet_names:
            st.error(f"Το sheet '{sheet_name}' δεν βρέθηκε. Διαθέσιμα: {xfile.sheet_names}")
            return None
        return pd.read_excel(xfile, sheet_name=sheet_name, engine="openpyxl")
    except Exception as e:
        st.error(f"Δεν άνοιξε το αρχείο: {e}")
        return None

# ───────────────────────────── SIDEBAR ─────────────────────────────
debug_mode = st.sidebar.toggle("🛠 Debug mode", value=True)
test_mode  = st.sidebar.toggle("🧪 Test mode (limit rows=50)", value=True)

st.sidebar.header("⚙️ BEX")
bex_mode = st.sidebar.radio("Πηγή BEX", ["Στήλη στο Excel", "Λίστα (comma-separated)"], index=0)
bex_list = set()
if bex_mode == "Λίστα (comma-separated)":
    bex_input = st.sidebar.text_area("BEX stores", "ESC01,FKM01,LND01,DRZ01,PKK01")
    bex_list = set(s.strip().upper() for s in bex_input.split(",") if s.strip())

st.sidebar.subheader("📄 Templates (.docx)")
tpl_bex    = st.sidebar.file_uploader("BEX template", type=["docx"])
tpl_nonbex = st.sidebar.file_uploader("Non-BEX template", type=["docx"])
st.sidebar.caption(
    "Placeholders: [[title]], [[store]], [[mobile_actual]], [[mobile_target]], "
    "[[fixed_actual]], [[fixed_target]], [[pending_mobile]], [[pending_fixed]], [[plan_vs_target]]"
)

# ───────────────────────────── MAIN INPUTS ─────────────────────────────
st.markdown("### 1) Ανέβασε Excel/CSV")
xls = st.file_uploader("Excel/CSV", type=["xlsx", "csv"])
sheet_name = st.text_input("Όνομα φύλλου (Sheet - μόνο για Excel)", value="Sheet1")

run = st.button("🔧 Generate")

# ───────────────────────────── MAIN ─────────────────────────────
if run:
    # Αρχικοί έλεγχοι
    if not xls:
        st.error("Ανέβασε αρχείο Excel ή CSV πρώτα.")
        st.stop()
    if not tpl_bex or not tpl_nonbex:
        st.error("Ανέβασε και τα δύο templates (.docx).")
        st.stop()

    st.info(
        f"📄 Δεδομένα: {len(xls.getbuffer())/1024:.1f} KB | "
        f"BEX tpl: {tpl_bex.size/1024:.1f} KB | Non-BEX tpl: {tpl_nonbex.size/1024:.1f} KB"
    )

    # Διαβάζουμε δεδομένα
    df = read_data(xls, sheet_name)
    if df is None or df.empty:
        st.error("Δεν βρέθηκαν δεδομένα στο αρχείο.")
        st.stop()

    st.success(f"OK: {len(df)} γραμμές, {len(df.columns)} στήλες.")
    import pandas as pd

# Φόρτωση Excel για να δούμε τις κεφαλίδες
xls_path = "sheet1.xlsx"

xfile = pd.ExcelFile(xls_path, engine="openpyxl")
print("📑 Sheets:", xfile.sheet_names)

df = pd.read_excel(xfile, sheet_name=xfile.sheet_names[0])
print("🔍 Headers:")
print(list(df.columns))

 if debug_mode:
        st.dataframe(df.head(10))

    cols = list(df.columns)

    # Auto-map βασισμένο στα headers
    col_store = pick(cols, "Shop Code", "Shop_Code", "ShopCode", "Shop code", "STORE", "Κατάστημα",
                 "shop", "store", "code καταστήματος", "ΚΩΔΙΚΟΣ ΚΑΤΑΣΤΗΜΑΤΟΣ", r"shop.?code")
    col_bex      = pick(cols, "BEX store", "BEX", r"bex.?store")
    col_mob_act  = pick(cols, "mobile actual", r"mobile.*actual")
    col_mob_tgt  = pick(cols, "mobile target", r"mobile.*target", "mobile plan")
    col_fix_tgt  = pick(cols, "target fixed", r"fixed.*target", "fixed plan total", "fixed plan")
    col_fix_act  = pick(cols, "total fixed", r"(total|sum).?fixed.*actual", "fixed actual")
    col_pend_mob = pick(cols, "TOTAL PENDING MOBILE", r"pending.*mobile")
    col_pend_fix = pick(cols, "TOTAL PENDING FIXED", r"pending.*fixed")
    col_plan_vs  = pick(cols, "plan vs target", r"plan.*vs.*target")

    # Εμφάνιση mapping
    with st.expander("Χαρτογράφηση (auto)"):
        st.write({
            "STORE": col_store, "BEX": col_bex,
            "mobile_actual": col_mob_act, "mobile_target": col_mob_tgt,
            "fixed_target": col_fix_tgt, "fixed_actual": col_fix_act,
            "pending_mobile": col_pend_mob, "pending_fixed": col_pend_fix,
            "plan_vs_target": col_plan_vs
        })

    if not col_store:
        st.error("Δεν βρέθηκε στήλη STORE (π.χ. 'Shop Code'). Διόρθωσε την κεφαλίδα ή πρόσθεσε alias.")
        st.stop()

    tpl_bex_bytes    = tpl_bex.read()
    tpl_nonbex_bytes = tpl_nonbex.read()

    # Out ZIP
    out_zip = io.BytesIO()
    zf = zipfile.ZipFile(out_zip, "w", zipfile.ZIP_DEFLATED)
    built = 0

    pbar = st.progress(0, text="Δημιουργία εγγράφων…")
    total = len(df) if not test_mode else min(50, len(df))

    for i, (_, row) in enumerate(df.iterrows(), start=1):
        if test_mode and i > total:
            st.info(f"🧪 Test mode: σταμάτησα στις {total} γραμμές.")
            break

        try:
            store = str(cell(row, col_store)).strip()
            if not store:
                pbar.progress(min(i / (total or 1), 1.0), text=f"Παράλειψη γραμμής {i} (κενό store)")
                continue

            store_up = store.upper()

            # BEX flag
            if bex_mode == "Λίστα (comma-separated)":
                is_bex = store_up in bex_list
            else:
                bex_val = str(cell(row, col_bex)).strip().lower()
                is_bex = bex_val in ("yes", "y", "1", "true", "ναι")

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

            doc = Document(io.BytesIO(tpl_bex_bytes if is_bex else tpl_nonbex_bytes))
            set_default_font(doc, "Aptos")
            replace_placeholders(doc, mapping)

            out_name = f"{store_up}_ReviewSep_PlanOct.docx"
            buf = io.BytesIO()
            doc.save(buf)
            zf.writestr(out_name, buf.getvalue())
            built += 1

            pbar.progress(min(i / (total or 1), 1.0), text=f"Φτιάχνω: {out_name} ({min(i, total)}/{total})")

        except Exception as e:
            st.warning(f"⚠️ Σφάλμα στη γραμμή {i}: {e}")
            if debug_mode:
                st.exception(e)

    zf.close()
    pbar.empty()

    if built == 0:
        st.error("Δεν δημιουργήθηκε αρχείο. Έλεγξε STORE mapping & templates.")
    else:
        st.success(f"Έτοιμα {built} αρχεία: {built}")
        st.download_button("⬇️ Κατέβασε ZIP", data=out_zip.getvalue(), file_name="reviews_from_excel.zip")
