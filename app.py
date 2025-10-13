# app.py
# Streamlit: Excel (xlsx) -> (BEX / Non-BEX) Review-Plan .docx (ZIP)
# Διαβάζει ΔΥΝΑΜΙΚΑ τιμές από γράμματα στηλών (A, N, O, P, Q, R, S, T, U, V, X, Y, AA, AB, AF, AH)

import io, zipfile, re
from typing import Any, Dict, Optional

import streamlit as st
from openpyxl import load_workbook
from docx import Document
from docx.oxml.ns import qn

# ───────── UI βασικά ─────────
st.set_page_config(page_title="Excel → Review/Plan Generator", layout="wide")
st.title("📊 Excel/CSV → 📄 Review/Plan Generator (BEX & Non-BEX)")

# ───────── helpers ─────────
def set_default_font(doc: Document, font_name: str = "Aptos") -> None:
    for style in doc.styles:
        if hasattr(style, "font"):
            try:
                style.font.name = font_name
                style._element.rPr.rFonts.set(qn("w:eastAsia"), font_name)
                style._element.rPr.rFonts.set(qn("w:cs"), font_name)
            except Exception:
                pass

def repl_placeholders(doc: Document, mapping: Dict[str, Any]) -> None:
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

def fmt(v: Any) -> Any:
    """Μικρό formatting: αριθμοί χωρίς .0, dates αφήνονται ως έχουν."""
    try:
        from datetime import datetime, date
        if isinstance(v, (int,)):
            return v
        if isinstance(v, float):
            if v.is_integer():
                return int(v)
            return round(v, 4)
        if isinstance(v, (datetime, date)):
            return v
    except Exception:
        pass
    return v if v is not None else ""

def try_get(ws, col_letter: Optional[str], row_idx: int):
    if not col_letter:
        return ""
    try:
        return ws[f"{col_letter.upper()}{row_idx}"].value
    except Exception:
        return ""

def find_store_letter(ws) -> Optional[str]:
    """Ψάχνει στην 1η γραμμή για 'Dealer_Code'/'Shop Code'/κ.λπ. και επιστρέφει γράμμα στήλης."""
    header_row = 1
    wanted = [r"dealer[_\s]*code", r"shop[_\s]*code", r"store", r"κατάστημα"]
    max_col = ws.max_column
    for c in range(1, max_col + 1):
        v = ws.cell(row=header_row, column=c).value
        if not v:
            continue
        vs = str(v).strip().lower()
        for pat in wanted:
            if re.search(pat, vs):
                # μετατροπή index -> γράμμα
                from openpyxl.utils import get_column_letter
                return get_column_letter(c)
    return None

def get_headers_row(ws, max_show=50):
    headers = []
    for c in range(1, ws.max_column + 1):
        headers.append(ws.cell(row=1, column=c).value)
        if len(headers) >= max_show:
            break
    return headers

# ───────── Sidebar / ρυθμίσεις ─────────
with st.sidebar:
    st.header("Ρυθμίσεις")
    debug = st.toggle("🛠 Debug mode", value=True)
    test_mode = st.toggle("🧪 Test mode (πρώτες 50 γραμμές)", value=True)

    st.subheader("Templates (.docx)")
    tpl_bex = st.file_uploader("BEX template", type=["docx"])
    tpl_nonbex = st.file_uploader("Non-BEX template", type=["docx"])
    st.caption("Placeholders: [[title]], [[plan_month]], [[store]], [[bex]], "
               "[[plan_vs_target]], [[mobile_actual]], [[mobile_target]], "
               "[[fixed_target]], [[fixed_actual]], [[voice_vs_target]], [[fixed_vs_target]], "
               "[[llu_actual]], [[nga_actual]], [[ftth_actual]], [[eon_tv_actual]], [[fwa_actual]], "
               "[[mobile_upgrades]], [[fixed_upgrades]], [[pending_mobile]], [[pending_fixed]]")

st.markdown("### 1) Ανέβασε Excel (xlsx)")
xls = st.file_uploader("Drag and drop file here", type=["xlsx"])
sheet_name = st.text_input("Όνομα φύλλου (Sheet - μόνο για Excel)", value="Sheet1")

st.markdown("### 🔧 Ρύθμιση γραμμών (headers & δεδομένα)")
col1, col2 = st.columns(2)
with col1:
    data_start_row = st.number_input("Πρώτη γραμμή δεδομένων", min_value=2, value=2, step=1)
with col2:
    sample_preview_rows = st.number_input("Προεπισκόπηση: πόσες πρώτες γραμμές να δείξω", min_value=1, value=1, step=1)

st.markdown("### 🏪 STORE & BEX")
bex_mode = st.radio("Πώς βρίσκουμε αν είναι BEX:", ["Σταθερή λίστα (DRZ01, …)", "Από στήλη (YES/NO)"], index=0)
bex_list_input = st.text_input("BEX/Non-BEX λίστα (comma-separated)", value="DRZ01,FKM01,ESC01,LND01,PKK01")
bex_letter = st.text_input("BEX Στήλη (YES/NO) – γράμμα", value="")

st.markdown("### ✉️ Mapping με γράμματα Excel (A, N, O, P, Q, R, S, T, U, V, X, Y, AA, AB, AF, AH)")
map_cols = {
    "plan_vs_target": st.text_input("plan vs target", value="A"),
    "mobile_actual": st.text_input("mobile actual", value="N"),
    "mobile_target": st.text_input("mobile target", value="O"),
    "fixed_target": st.text_input("fixed target", value="P"),
    "fixed_actual": st.text_input("total fixed actual", value="Q"),
    "voice_vs_target": st.text_input("voice Vs target", value="R"),
    "fixed_vs_target": st.text_input("fixed vs target", value="S"),
    "llu_actual": st.text_input("llu actual", value="T"),
    "nga_actual": st.text_input("nga actual", value="U"),
    "ftth_actual": st.text_input("ftth actual", value="V"),
    "eon_tv_actual": st.text_input("eon tv actual", value="X"),
    "fwa_actual": st.text_input("fwa actual", value="Y"),
    "mobile_upgrades": st.text_input("mobile upgrades", value="AA"),
    "fixed_upgrades": st.text_input("fixed upgrades", value="AB"),
    "pending_mobile": st.text_input("total pending mobile", value="AF"),
    "pending_fixed": st.text_input("total pending fixed", value="AH"),
}

st.markdown("### 🗓️ Τίτλοι")
title_month = st.text_input("Title (π.χ. 'Review September 2025 — Plan October 2025')",
                            value="Review September 2025 — Plan October 2025")
plan_month = st.text_input("Plan month πεδίο (αν το θες ξεχωριστά)", value="Review September 2025 — Plan October 2025")

run = st.button("🔧 Generate")

# ───────── Main logic ─────────
if run:
    # Έλεγχοι
    if not xls:
        st.error("Ανέβασε Excel (.xlsx) πρώτα.")
        st.stop()
    if not tpl_bex or not tpl_nonbex:
        st.error("Ανέβασε και τα δύο templates (.docx).")
        st.stop()

    try:
        wb = load_workbook(filename=xls, data_only=True)
        if sheet_name not in wb.sheetnames:
            st.error(f"Το sheet '{sheet_name}' δεν βρέθηκε. Διαθέσιμα: {wb.sheetnames}")
            st.stop()
        ws = wb[sheet_name]
    except Exception as e:
        st.error(f"Αποτυχία φόρτωσης Excel: {e}")
        st.stop()

    # Auto-detect γράμμα STORE (από header) + δυνατότητα override
    auto_store_letter = find_store_letter(ws)
    st.info(f"🔎 Βρήκα STORE στήλη: {auto_store_letter or '—'} (από header row 1)")
    store_letter = st.text_input("STORE letter (αν θέλεις override)", value=auto_store_letter or "")

    # Δείξε Headers (πρώτες 50 στήλες)
    if debug:
        st.write("**Headers όπως διαβάζονται (row 1):**", get_headers_row(ws, max_show=50))

    # Preview 2ης γραμμής (ή όσες ζήτησες)
    st.markdown("### 🔍 Mapping preview (από 2η γραμμή)")
    bex_set = set(s.strip().upper() for s in bex_list_input.split(",") if s.strip())

    previews = []
    for r in range(data_start_row, data_start_row + int(sample_preview_rows)):
        store_val = fmt(try_get(ws, store_letter, r))
        sample = {"row_excel": r, "store": {"from": "header", "value": store_val}}
        for key, letter in map_cols.items():
            sample[key] = fmt(try_get(ws, letter, r))
        # προσδιορισμός BEX
        if bex_mode.startswith("Σταθερή"):
            sample["bex"] = "YES" if str(store_val).upper() in bex_set else "NO"
        else:
            raw = str(fmt(try_get(ws, bex_letter, r))).strip().lower()
            sample["bex"] = "YES" if raw in ("yes", "y", "1", "true", "ναι") else "NO"
        previews.append(sample)
    st.write(previews)

    # Δημιουργία αρχείων
    tpl_bex_bytes = tpl_bex.read()
    tpl_nonbex_bytes = tpl_nonbex.read()

    out_zip = io.BytesIO()
    zf = zipfile.ZipFile(out_zip, "w", zipfile.ZIP_DEFLATED)
    built = 0

    max_rows = ws.max_row
    last_row = max_rows if not test_mode else min(max_rows, data_start_row + 49)
    pbar = st.progress(0.0, text="Δημιουργία εγγράφων…")

    for i, r in enumerate(range(data_start_row, last_row + 1), start=1):
        store_val = str(fmt(try_get(ws, store_letter, r))).strip()
        if not store_val:
            pbar.progress(min(i / max(1, (last_row - data_start_row + 1)), 1.0),
                          text=f"Παράλειψη γραμμής {r} (κενό STORE)")
            continue

        row_vals = {k: fmt(try_get(ws, letter, r)) for k, letter in map_cols.items()}

        # BEX
        if bex_mode.startswith("Σταθερή"):
            is_bex = store_val.upper() in bex_set
        else:
            raw = str(fmt(try_get(ws, bex_letter, r))).strip().lower()
            is_bex = raw in ("yes", "y", "1", "true", "ναι")

        # Word mapping
        mapping = {
            "title": title_month,
            "plan_month": plan_month,
            "store": store_val.upper(),
            "bex": "YES" if is_bex else "NO",
            **row_vals
        }

        try:
            doc = Document(io.BytesIO(tpl_bex_bytes if is_bex else tpl_nonbex_bytes))
            set_default_font(doc, "Aptos")
            repl_placeholders(doc, mapping)

            out_name = f"{store_val.upper()}_Review_Plan.docx"
            buf = io.BytesIO()
            doc.save(buf)
            zf.writestr(out_name, buf.getvalue())
            built += 1
            pbar.progress(min(i / max(1, (last_row - data_start_row + 1)), 1.0),
                          text=f"Φτιάχνω: {out_name} ({i}/{last_row - data_start_row + 1})")
        except Exception as e:
            st.warning(f"⚠️ Σφάλμα στη γραμμή {r}: {e}")

    zf.close()
    pbar.empty()

    if built == 0:
        st.error("Δεν δημιουργήθηκε αρχείο. Έλεγξε STORE mapping & templates.")
    else:
        st.success(f"Έτοιμα {built} αρχεία.")
        st.download_button("⬇️ Κατέβασε ZIP", data=out_zip.getvalue(),
                           file_name="reviews_from_excel.zip", mime="application/zip")