# app.py — Excel/CSV → BEX & Non-BEX Review/Plan Generator (letter-true mapping)

import io, re, zipfile
from typing import Any, Dict

import pandas as pd
import streamlit as st
from docx import Document
from docx.oxml.ns import qn

# ---- ΝΕΟ: δουλεύουμε με απόλυτα γράμματα Excel ----
try:
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter, column_index_from_string
except Exception:
    load_workbook = None

# ── UI ───────────────────────────────────────────────
st.set_page_config(page_title="Excel → Review/Plan Generator", layout="wide")
st.title("📊 Excel/CSV → 📄 Review/Plan Generator (BEX & Non-BEX)")

# ── helpers ─────────────────────────────────────────
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
    pat = re.compile(r"\[\[([A-Za-z0-9_]+)\]\]")
    def subfun(s: str) -> str:
        return pat.sub(lambda m: "" if mapping.get(m.group(1)) is None else str(mapping.get(m.group(1), "")), s)
    for p in doc.paragraphs:
        for r in p.runs:
            r.text = subfun(r.text)
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        r.text = subfun(r.text)

def norm_letter(s: str) -> str:
    return re.sub(r"[^A-Za-z]", "", str(s)).upper()

def letter_to_index(letter: str) -> int | None:
    s = norm_letter(letter)
    if not s:
        return None
    try:
        return column_index_from_string(s)  # 1-based
    except Exception:
        return None

# ---- Διαβάζουμε και ως DataFrame (για preview) και ως openpyxl (για letters)
def read_excel_both(xls, sheet_name: str):
    """Επιστρέφει (df, wb, ws). Αν είναι CSV: (df, None, None)."""
    try:
        fname = getattr(xls, "name", "")
        if fname.lower().endswith(".csv"):
            df = pd.read_csv(xls)
            df = df.rename(columns=lambda c: str(c).strip())
            return df, None, None
        # xlsx
        if load_workbook is None:
            st.error("Λείπει το openpyxl. Πρόσθεσέ το στο requirements.txt")
            return None, None, None
        # 1) openpyxl
        xls.seek(0)
        wb = load_workbook(filename=xls, data_only=True, read_only=True)
        if sheet_name not in wb.sheetnames:
            st.error(f"Το sheet '{sheet_name}' δεν βρέθηκε. Διαθέσιμα: {wb.sheetnames}")
            return None, None, None
        ws = wb[sheet_name]
        # 2) pandas (προαιρετικά για γρήγορο head/preview)
        # διαβάζουμε ξανά γιατί load_workbook κατανάλωσε το stream
        xls.seek(0)
        df = pd.read_excel(xls, sheet_name=sheet_name, engine="openpyxl")
        df = df.rename(columns=lambda c: str(c).strip())
        return df, wb, ws
    except Exception as e:
        st.error(f"Δεν άνοιξε το αρχείο: {e}")
        return None, None, None

def ws_header_by_letter(ws, letter: str) -> str | None:
    """Header από τη ΓΡΑΜΜΗ 1 για συγκεκριμένο γράμμα (A=1)."""
    idx = letter_to_index(letter)
    if not idx:
        return None
    cell = ws.cell(row=1, column=idx)
    val = cell.value
    return None if val is None else str(val).strip()

def ws_value(ws, data_row_1based: int, letter: str):
    """Τιμή από συγκεκριμένη γραμμή/στήλη. data_row_1based: 2=πρώτη σειρά δεδομένων."""
    idx = letter_to_index(letter)
    if not idx:
        return ""
    v = ws.cell(row=data_row_1based, column=idx).value
    return "" if v is None else v

# ── SIDEBAR ─────────────────────────────────────────
debug_mode = st.sidebar.toggle("🛠 Debug mode", value=True)
test_mode  = st.sidebar.toggle("🧪 Test mode (limit rows=50)", value=True)

st.sidebar.header("⚙️ BEX")
BEX_STORES = {"DRZ01", "FKM01", "ESC01", "LND01", "PKK01"}
st.sidebar.info("BEX stores: DRZ01, FKM01, ESC01, LND01, PKK01")

st.sidebar.subheader("📄 Templates (.docx)")
tpl_bex    = st.sidebar.file_uploader("BEX template", type=["docx"])
tpl_nonbex = st.sidebar.file_uploader("Non-BEX template", type=["docx"])
st.sidebar.caption("Placeholders: [[title]], [[store]], [[plan_vs_target]], [[mobile_actual]], [[mobile_target]], [[fixed_actual]], [[fixed_target]], [[pending_mobile]], [[pending_fixed]], κ.ά.")

st.sidebar.subheader("📎 Manual mapping (Excel letters)")
# Δώσε Ο,ΤΙ γράμματα θες — τώρα είναι απόλυτα στο φύλλο
letter_store      = st.sidebar.text_input("STORE (π.χ. Dealer_Code)", "A")   # ← στη δική σου περίπτωση είναι G, βάλε G
letter_plan_vs    = st.sidebar.text_input("plan vs target", "B")
letter_mobile_act = st.sidebar.text_input("mobile actual", "C")
letter_mobile_tgt = st.sidebar.text_input("mobile target", "D")
letter_fixed_tgt  = st.sidebar.text_input("fixed target", "E")
letter_fixed_act  = st.sidebar.text_input("total fixed actual", "F")
letter_voice_vs   = st.sidebar.text_input("voice vs target", "G")
letter_fixed_vs   = st.sidebar.text_input("fixed vs target", "H")
letter_llu        = st.sidebar.text_input("llu actual", "T")
letter_nga        = st.sidebar.text_input("nga actual", "U")
letter_ftth       = st.sidebar.text_input("ftth actual", "V")
letter_eon        = st.sidebar.text_input("eon tv actual", "X")
letter_fwa        = st.sidebar.text_input("fwa actual", "Y")
letter_mob_upg    = st.sidebar.text_input("mobile upgrades", "AA")
letter_fix_upg    = st.sidebar.text_input("fixed upgrades", "AB")
letter_pend_mob   = st.sidebar.text_input("total pending mobile", "AF")
letter_pend_fix   = st.sidebar.text_input("total pending fixed", "AH")

LETTERS = {
    "store_letter":     letter_store,
    "plan_vs_target":   letter_plan_vs,
    "mobile_actual":    letter_mobile_act,
    "mobile_target":    letter_mobile_tgt,
    "fixed_target":     letter_fixed_tgt,
    "fixed_actual":     letter_fixed_act,
    "voice_vs_target":  letter_voice_vs,
    "fixed_vs_target":  letter_fixed_vs,
    "llu_actual":       letter_llu,
    "nga_actual":       letter_nga,
    "ftth_actual":      letter_ftth,
    "eon_tv_actual":    letter_eon,
    "fwa_actual":       letter_fwa,
    "mobile_upgrades":  letter_mob_upg,
    "fixed_upgrades":   letter_fix_upg,
    "pending_mobile":   letter_pend_mob,
    "pending_fixed":    letter_pend_fix,
}

# ── MAIN INPUTS ─────────────────────────────────────
st.markdown("### 1) Ανέβασε Excel/CSV")
xls = st.file_uploader("Excel/CSV", type=["xlsx", "csv"])
sheet_name = st.text_input("Όνομα φύλλου (Sheet - μόνο για Excel)", value="Sheet1")
run = st.button("🔧 Generate")

# ── MAIN ────────────────────────────────────────────
if run:
    if not xls:
        st.error("Ανέβασε αρχείο Excel ή CSV πρώτα.")
        st.stop()
    if not tpl_bex or not tpl_nonbex:
        st.error("Ανέβασε και τα δύο templates.")
        st.stop()

    df, wb, ws = read_excel_both(xls, sheet_name)
    if df is None:
        st.stop()

    st.success(f"OK: {len(df)} γραμμές (DF preview), {len(df.columns)} στήλες DF.")
    if debug_mode and not df.empty:
        st.dataframe(df.head(10))

    # Preview: γράμμα → header (row1) → τιμή στην πρώτη γραμμή δεδομένων (row2)
    preview = {}
    if ws is not None:
        data_start_row = 2  # υποθέτουμε headers στη row 1
        for key, L in LETTERS.items():
            hdr = ws_header_by_letter(ws, L)
            sample = ws_value(ws, data_start_row, L)
            preview[key] = {"letter": norm_letter(L), "header_row1": hdr, "sample_row2": sample}
        with st.expander("🧭 Letters → Header(Row1) → Sample(Row2)"):
            st.json(preview)
    else:
        # CSV fallback: A=1η στήλη DF, B=2η κ.ο.κ.
        for key, L in LETTERS.items():
            idx = letter_to_index(L)
            if idx and 0 < idx <= len(df.columns):
                hdr = df.columns[idx-1]
                sample = df.iloc[0, idx-1] if len(df) else ""
            else:
                hdr = None
                sample = None
            preview[key] = {"letter": norm_letter(L), "header_row1": hdr, "sample_row2": sample}
        with st.expander("🧭 Letters (CSV mode)"):
            st.json(preview)

    tpl_bex_bytes = tpl_bex.read()
    tpl_nonbex_bytes = tpl_nonbex.read()

    out_zip = io.BytesIO()
    zf = zipfile.ZipFile(out_zip, "w", zipfile.ZIP_DEFLATED)
    built = 0

    # Πόσες γραμμές; Από το worksheet για ακριβές γράμμα-βασισμένο διαβασμα
    if ws is not None:
        max_rows = ws.max_row
        total = (max_rows - 1) if not test_mode else min(50, max_rows - 1)
        pbar = st.progress(0, text="Δημιουργία εγγράφων…")

        for i, excel_row in enumerate(range(2, max_rows + 1), start=1):
            if test_mode and i > total:
                break
            try:
                store = str(ws_value(ws, excel_row, letter_store)).strip()
                if not store:
                    pbar.progress(min(i/(total or 1), 1.0), text=f"Skip row {excel_row} (empty store)")
                    continue

                is_bex = store.upper() in BEX_STORES

                mapping = {
                    "title": f"Review September 2025 — Plan October 2025 — {store}",
                    "store": store,
                }
                for key, L in LETTERS.items():
                    if key == "store_letter":
                        continue
                    mapping[key] = ws_value(ws, excel_row, L)

                doc = Document(io.BytesIO(tpl_bex_bytes if is_bex else tpl_nonbex_bytes))
                set_default_font(doc, "Aptos")
                replace_placeholders(doc, mapping)

                buf = io.BytesIO()
                doc.save(buf)
                zf.writestr(f"{store}_ReviewSep_PlanOct.docx", buf.getvalue())
                built += 1
                pbar.progress(min(i/(total or 1), 1.0), text=f"Φτιάχνω: {store} ({min(i,total)}/{total})")
            except Exception as e:
                st.warning(f"⚠️ Σφάλμα στη γραμμή Excel {excel_row}: {e}")
                if debug_mode:
                    st.exception(e)

        pbar.empty()

    else:
        # CSV fallback: γράμματα → θέσεις DF
        total = len(df) if not test_mode else min(50, len(df))
        pbar = st.progress(0, text="Δημιουργία εγγράφων (CSV)…")
        for i, (_, row) in enumerate(df.iterrows(), start=1):
            if test_mode and i > total:
                break
            try:
                idx_store = letter_to_index(letter_store)
                store = str(row.iloc[idx_store-1]).strip() if idx_store and idx_store-1 < len(df.columns) else ""
                if not store:
                    continue
                is_bex = store.upper() in BEX_STORES

                mapping = {
                    "title": f"Review September 2025 — Plan October 2025 — {store}",
                    "store": store,
                }
                for key, L in LETTERS.items():
                    if key == "store_letter":
                        continue
                    idx = letter_to_index(L)
                    mapping[key] = row.iloc[idx-1] if idx and idx-1 < len(df.columns) else ""

                doc = Document(io.BytesIO(tpl_bex_bytes if is_bex else tpl_nonbex_bytes))
                set_default_font(doc, "Aptos")
                replace_placeholders(doc, mapping)

                buf = io.BytesIO()
                doc.save(buf)
                zf.writestr(f"{store}_ReviewSep_PlanOct.docx", buf.getvalue())
                built += 1
                pbar.progress(min(i/(total or 1), 1.0), text=f"Φτιάχνω: {store} ({min(i,total)}/{total})")
            except Exception as e:
                st.warning(f"⚠️ Σφάλμα στη γραμμή {i}: {e}")
                if debug_mode:
                    st.exception(e)
        pbar.empty()

    zf.close()
    if built == 0:
        st.error("Δεν δημιουργήθηκε αρχείο.")
    else:
        st.success(f"✅ Έτοιμα {built} αρχεία.")
        st.download_button("⬇️ Κατέβασε ZIP", data=out_zip.getvalue(), file_name="reviews_from_excel.zip")