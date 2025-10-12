# app.py — Streamlit: Excel/CSV -> .docx (BEX/Non-BEX) Generator

import io
import re
import zipfile
from typing import Any, Dict, Optional

import pandas as pd
import streamlit as st
from docx import Document
from docx.oxml.ns import qn

# ──────────────────────────────────────────────────────────────────────────────
# UI CONFIG
# ──────────────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="Excel → Review/Plan Generator", layout="wide")
st.title("📊 Excel/CSV → 📄 Review/Plan Generator (BEX & Non-BEX)")

DEBUG = st.sidebar.toggle("🛠 Debug mode", value=True)
TEST_MODE = st.sidebar.toggle("🧪 Test mode (limit rows=50)", value=False)

# ──────────────────────────────────────────────────────────────────────────────
# CONSTANTS
# ──────────────────────────────────────────────────────────────────────────────
DEFAULT_BEX_STORES = {"DRZ01", "FKM01", "ESC01", "LND01", "PKK01"}

PLACEHOLDERS = [
    "title", "store", "plan_month", "plan_vs_target",
    "mobile_actual", "mobile_target",
    "fixed_actual", "fixed_target",
    "voice_vs_target", "fixed_vs_target",
    "llu_actual", "nga_actual", "ftth_actual",
    "eon_tv_actual", "fwa_actual",
    "mobile_upgrades", "fixed_upgrades",
    "pending_mobile", "pending_fixed",
    "bex",
]

# ──────────────────────────────────────────────────────────────────────────────
# HELPERS
# ──────────────────────────────────────────────────────────────────────────────
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
        key = lambda m: m.group(1)
        return pattern.sub(lambda m: "" if mapping.get(key(m)) is None else str(mapping.get(key(m), "")), s)

    for p in doc.paragraphs:
        for r in p.runs:
            r.text = subfun(r.text)
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        r.text = subfun(r.text)

def excel_letter_to_index(letter: str) -> Optional[int]:
    """
    Μετατρέπει γράμμα(τα) Excel -> 0-based index για pandas (A->0, B->1, ... Z->25, AA->26, ...).
    Δέχεται επίσης κενά/πεζά. Επιστρέφει None αν είναι άδειρο.
    """
    if not letter:
        return None
    s = letter.strip().upper()
    if not re.fullmatch(r"[A-Z]+", s):
        return None
    n = 0
    for ch in s:
        n = n * 26 + (ord(ch) - ord('A') + 1)
    return n - 1

def safe_get_by_letter(row: pd.Series, df: pd.DataFrame, letter: str):
    """Παίρνει τιμή από γραμμή row με βάση γράμμα Excel (0-based index πάνω στη σειρά των στηλών του df)."""
    idx = excel_letter_to_index(letter)
    if idx is None:
        return ""
    if 0 <= idx < df.shape[1]:
        val = row.iloc[idx]
        return "" if pd.isna(val) else val
    return ""

def read_table(file, sheet_name: str) -> pd.DataFrame | None:
    """Δέχεται .xlsx ή .csv. Διαβάζει όλο το sheet/csv με 1η γραμμή headers."""
    try:
        name = getattr(file, "name", "").lower()
        if name.endswith(".csv"):
            df = pd.read_csv(file)
            st.write("📑 Sheets:", ["CSV Data"])
            return df
        # XLSX
        xf = pd.ExcelFile(file, engine="openpyxl")
        st.write("📑 Sheets:", xf.sheet_names)
        if sheet_name not in xf.sheet_names:
            st.error(f"Το sheet '{sheet_name}' δεν βρέθηκε. Διαθέσιμα: {xf.sheet_names}")
            return None
        df = pd.read_excel(xf, sheet_name=sheet_name, engine="openpyxl")
        return df
    except Exception as e:
        st.error(f"Δεν άνοιξε το αρχείο: {e}")
        return None

# ──────────────────────────────────────────────────────────────────────────────
# TEMPLATES (DOCX)
# ──────────────────────────────────────────────────────────────────────────────
st.sidebar.subheader("📄 Templates (.docx)")
tpl_bex_file    = st.sidebar.file_uploader("BEX template", type=["docx"])
tpl_nonbex_file = st.sidebar.file_uploader("Non-BEX template", type=["docx"])
st.sidebar.caption(
    "Χρησιμοποίησε placeholders στο Word: "
    + ", ".join(f"[[{k}]]" for k in PLACEHOLDERS)
)

# ──────────────────────────────────────────────────────────────────────────────
# BEX MODE
# ──────────────────────────────────────────────────────────────────────────────
st.sidebar.subheader("🏷️ BEX detection")
bex_mode = st.sidebar.radio(
    "Πώς βρίσκουμε αν είναι BEX;",
    ["Σταθερή λίστα (DRZ01, FKM01, ESC01, LND01, PKK01)", "Από στήλη (YES/NO)"],
    index=0,
)
bex_yesno_letter = ""
if bex_mode == "Από στήλη (YES/NO)":
    bex_yesno_letter = st.sidebar.text_input("Γράμμα στήλης με YES/NO (π.χ. J)", value="")

# ──────────────────────────────────────────────────────────────────────────────
# INPUT DATA
# ──────────────────────────────────────────────────────────────────────────────
st.markdown("### 1) Ανέβασε Excel/CSV")
data_file = st.file_uploader("Excel/CSV", type=["xlsx", "csv"])
sheet_name = st.text_input("Όνομα φύλλου (Sheet - μόνο για Excel)", value="Sheet1")

# ──────────────────────────────────────────────────────────────────────────────
# MANUAL MAPPING (letters) & STORE column (by header)
# ──────────────────────────────────────────────────────────────────────────────
st.markdown("### 2) Mapping στηλών")
col1, col2 = st.columns([1, 1])

with col1:
    st.write("**Manual mapping με γράμματα Excel** (A, B, …, Z, AA, AB, …)")
    letter_plan_vs   = st.text_input("plan vs target", value="")
    letter_mob_act   = st.text_input("mobile actual", value="")
    letter_mob_tgt   = st.text_input("mobile target", value="")
    letter_fix_tgt   = st.text_input("fixed target", value="")
    letter_fix_act   = st.text_input("total fixed actual", value="")
    letter_voice_pct = st.text_input("voice Vs target (%)", value="")
    letter_fixed_pct = st.text_input("fixed Vs target (%)", value="")
    letter_llu       = st.text_input("llu actual", value="")
    letter_nga       = st.text_input("nga actual", value="")
    letter_ftth      = st.text_input("ftth actual", value="")
    letter_eon       = st.text_input("eon tv actual", value="")
    letter_fwa       = st.text_input("fwa actual", value="")
    letter_mob_upg   = st.text_input("mobile upgrades", value="")
    letter_fix_upg   = st.text_input("fixed upgrades", value="")
    letter_pend_mob  = st.text_input("total pending mobile", value="")
    letter_pend_fix  = st.text_input("total pending fixed", value="")

with col2:
    st.write("**STORE column** (επίλεξε από headers)")
    # Θα γεμίσει αφού φορτωθεί το αρχείο
    store_col_placeholder = st.empty()
    plan_month = st.text_input("Κείμενο μήνα για τίτλο (π.χ. 'September 2025 / Plan October 2025')", value="September 2025 — Plan October 2025")

# ──────────────────────────────────────────────────────────────────────────────
# RUN
# ──────────────────────────────────────────────────────────────────────────────
run = st.button("🔧 Generate")

if run:
    # Έλεγχοι αρχείων
    if not data_file:
        st.error("Ανέβασε αρχείο Excel ή CSV πρώτα.")
        st.stop()
    if not tpl_bex_file or not tpl_nonbex_file:
        st.error("Ανέβασε και τα δύο templates (.docx).")
        st.stop()

    st.info(
        f"📄 Δεδομένα: {len(data_file.getbuffer())/1024:.1f} KB | "
        f"BEX tpl: {tpl_bex_file.size/1024:.1f} KB | Non-BEX tpl: {tpl_nonbex_file.size/1024:.1f} KB"
    )

    # Διαβάζουμε δεδομένα
    df = read_table(data_file, sheet_name)
    if df is None or df.empty:
        st.error("Δεν βρέθηκαν δεδομένα.")
        st.stop()

    st.success(f"OK: {len(df)} γραμμές, {len(df.columns)} στήλες.")
    if DEBUG:
        with st.expander("🔎 Headers όπως τους βλέπουμε"):
            st.write(list(df.columns))
        st.dataframe(df.head(10))

    # Επιλογή STORE μετά τη φόρτωση
    store_col = store_col_placeholder.selectbox("Στήλη STORE (header)", options=list(df.columns), index=0)

    # Προεπισκόπηση mapping: δείξε letter -> header(Row1) -> sample(Row2)
    with st.expander("🔤 Προεπισκόπηση mapping (Letters → Header(Row1) → Sample(Row2))"):
        preview = {}
        letters = [
            ("store_letter", None),  # το store θα έρθει από header επιλογής
            ("plan_vs_target", letter_plan_vs),
            ("mobile_actual", letter_mob_act),
            ("mobile_target", letter_mob_tgt),
            ("fixed_target", letter_fix_tgt),
            ("fixed_actual", letter_fix_act),
            ("voice_vs_target", letter_voice_pct),
            ("fixed_vs_target", letter_fixed_pct),
            ("llu_actual", letter_llu),
            ("nga_actual", letter_nga),
            ("ftth_actual", letter_ftth),
            ("eon_tv_actual", letter_eon),
            ("fwa_actual", letter_fwa),
            ("mobile_upgrades", letter_mob_upg),
            ("fixed_upgrades", letter_fix_upg),
            ("pending_mobile", letter_pend_mob),
            ("pending_fixed", letter_pend_fix),
        ]
        if len(df) >= 2:
            r0, r1 = df.iloc[0], df.iloc[1]
        else:
            r0 = r1 = df.iloc[0] if len(df) >= 1 else pd.Series(dtype=object)

        for key, letter in letters:
            if key == "store_letter":
                preview[key] = {"header": store_col, "sample_row2": ("" if df.empty else r1.get(store_col, ""))}
            else:
                idx = excel_letter_to_index(letter or "")
                header = df.columns[idx] if (idx is not None and 0 <= idx < df.shape[1]) else ""
                sample = (r1.iloc[idx] if (idx is not None and 0 <= idx < df.shape[1] and len(df) >= 2) else "")
                preview[key] = {"letter": (letter or ""), "header_row1": header, "sample_row2": sample}
        st.write(preview)

    # Διαβάζουμε templates σε μνήμη
    tpl_bex_bytes = tpl_bex_file.read()
    tpl_nonbex_bytes = tpl_nonbex_file.read()

    # ZIP out
    out_zip = io.BytesIO()
    zf = zipfile.ZipFile(out_zip, "w", zipfile.ZIP_DEFLATED)
    built = 0

    total_rows = len(df)
    if TEST_MODE:
        total_rows = min(total_rows, 50)

    pbar = st.progress(0, text="Δημιουργία εγγράφων…")

    for i, (_, row) in enumerate(df.iterrows(), start=1):
        if TEST_MODE and i > total_rows:
            break

        try:
            store_val = str(row.get(store_col, "")).strip()
            if not store_val:
                pbar.progress(min(i / total_rows, 1.0), text=f"Παράλειψη γραμμής {i} (κενό STORE)")
                continue

            # BEX flag
            if bex_mode.startswith("Σταθερή"):
                is_bex = store_val.upper() in DEFAULT_BEX_STORES
            else:
                bex_raw = safe_get_by_letter(row, df, bex_yesno_letter).strip().lower()
                is_bex = bex_raw in {"yes", "y", "1", "true", "ναι"}

            # Ανάγνωση από γράμματα (ό,τι είναι κενό, μένει κενό)
            mapping = {
                "title": f"Review {plan_month} — {store_val.upper()}",
                "store": store_val.upper(),
                "plan_month": plan_month,
                "bex": "YES" if is_bex else "NO",

                "plan_vs_target":  safe_get_by_letter(row, df, letter_plan_vs),
                "mobile_actual":   safe_get_by_letter(row, df, letter_mob_act),
                "mobile_target":   safe_get_by_letter(row, df, letter_mob_tgt),
                "fixed_target":    safe_get_by_letter(row, df, letter_fix_tgt),
                "fixed_actual":    safe_get_by_letter(row, df, letter_fix_act),
                "voice_vs_target": safe_get_by_letter(row, df, letter_voice_pct),
                "fixed_vs_target": safe_get_by_letter(row, df, letter_fixed_pct),
                "llu_actual":      safe_get_by_letter(row, df, letter_llu),
                "nga_actual":      safe_get_by_letter(row, df, letter_nga),
                "ftth_actual":     safe_get_by_letter(row, df, letter_ftth),
                "eon_tv_actual":   safe_get_by_letter(row, df, letter_eon),
                "fwa_actual":      safe_get_by_letter(row, df, letter_fwa),
                "mobile_upgrades": safe_get_by_letter(row, df, letter_mob_upg),
                "fixed_upgrades":  safe_get_by_letter(row, df, letter_fix_upg),
                "pending_mobile":  safe_get_by_letter(row, df, letter_pend_mob),
                "pending_fixed":   safe_get_by_letter(row, df, letter_pend_fix),
            }

            # Ποιο template;
            tpl_bytes = tpl_bex_bytes if is_bex else tpl_nonbex_bytes

            # Φτιάχνουμε docx
            doc = Document(io.BytesIO(tpl_bytes))
            set_default_font(doc, "Aptos")
            replace_placeholders(doc, mapping)

            # Αποθήκευση στο ZIP
            out_name = f"{store_val.upper()}_Review.docx"
            buf = io.BytesIO()
            doc.save(buf)
            zf.writestr(out_name, buf.getvalue())
            built += 1

            pbar.progress(min(i / total_rows, 1.0), text=f"Φτιάχνω: {out_name} ({min(i, total_rows)}/{total_rows})")

            if DEBUG and i <= 3:
                with st.expander(f"🧩 Mapping γραμμής {i} (preview)"):
                    st.write(mapping)

        except Exception as e:
            st.warning(f"⚠️ Σφάλμα στη γραμμή {i}: {e}")
            if DEBUG:
                st.exception(e)

    zf.close()
    pbar.empty()

    if built == 0:
        st.error("Δεν δημιουργήθηκε αρχείο. Έλεγξε STORE mapping, γράμματα και templates.")
    else:
        st.success(f"Έτοιμα {built} αρχεία.")
        st.download_button("⬇️ Κατέβασε ZIP", data=out_zip.getvalue(), file_name="reviews_from_excel.zip")