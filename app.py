# app.py
# Streamlit: Excel/CSV → (BEX / Non-BEX) Review-Plan .docx (ZIP)

import io, re, zipfile
from typing import Any, Dict, Optional

import pandas as pd
import streamlit as st
from docx import Document
from docx.oxml.ns import qn
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string

st.set_page_config(page_title="Excel → Review/Plan Generator", layout="wide")
st.title("📚 Excel/CSV → 📄 Review/Plan Generator (BEX & Non-BEX)")

# ---------------- Helpers ----------------
def set_default_font(doc: Document, font_name: str = "Aptos") -> None:
    for style in doc.styles:
        if hasattr(style, "font"):
            try:
                style.font.name = font_name
                style._element.rPr.rFonts.set(qn("w:eastAsia"), font_name)
                style._element.rPr.rFonts.set(qn("w:cs"), font_name)
            except Exception:
                pass

def replace_placeholders(doc: Document, mapping: Dict[str, Any]) -> None:
    pattern = re.compile(r"\[\[([A-Za-z0-9_]+)\]\]")
    def subfun(s: str) -> str:
        return pattern.sub(lambda m: "" if mapping.get(m.group(1)) is None else str(mapping.get(m.group(1), "")), s)
    for p in doc.paragraphs:
        for r in p.runs:
            r.text = subfun(r.text)
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        r.text = subfun(r.text)

def pick(columns, *aliases) -> str:
    # πιάνουμε store από header αν δεν δίνεται γράμμα
    def normkey(x: str) -> str:
        return re.sub(r"[\s\-_\.]+", "", str(x).strip().lower())
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

def safe(v):
    if v is None:
        return ""
    if isinstance(v, float) and pd.isna(v):
        return ""
    return v

# ---- OpenPyXL readers (για γράμματα στηλών) ----
def get_cell_value(ws, col_letter: Optional[str], row_idx_1based: int):
    """Διάβασε ακριβώς το κελί με openpyxl: (γράμμα, row 1-based)."""
    if not col_letter:
        return None
    try:
        col_idx = column_index_from_string(col_letter.strip().upper())
    except Exception:
        return None
    cell = ws.cell(row=row_idx_1based, column=col_idx)
    return cell.value

def get_value_by_letter(ws, letter: Optional[str], data_row_1based: int) -> str:
    v = get_cell_value(ws, letter, data_row_1based)
    if v is None:
        return ""
    # Φόρμες/ποσοστά/αριθμοί έρχονται ως value (data_only=True στο workbook)
    return str(v)

# ---------------- Sidebar ----------------
with st.sidebar:
    st.header("Ρυθμίσεις")
    debug_mode = st.toggle("🛠 Debug mode", value=True)
    test_mode  = st.toggle("🧪 Test mode (πρώτες 50 γραμμές)", value=True)

    st.subheader("Templates (.docx)")
    tpl_bex    = st.file_uploader("BEX template", type=["docx"], key="tpl_bex")
    tpl_nonbex = st.file_uploader("Non-BEX template", type=["docx"], key="tpl_nonbex")

    st.caption("Placeholders: [[title]] [[plan_month]] [[store]] [[bex]] "
               "[[plan_vs_target]] [[mobile_actual]] [[mobile_target]] [[fixed_target]] [[fixed_actual]] "
               "[[voice_vs_target]] [[fixed_vs_target]] [[llu_actual]] [[nga_actual]] [[ftth_actual]] "
               "[[eon_tv_actual]] [[fwa_actual]] [[mobile_upgrades]] [[fixed_upgrades]] "
               "[[pending_mobile]] [[pending_fixed]]")

# ---------------- Main inputs ----------------
st.markdown("### 1) Ανέβασε Excel/CSV")
xls = st.file_uploader("Excel/CSV", type=["xlsx", "csv"], key="xls")
sheet_name = st.text_input("Όνομα φύλλου (Sheet - μόνο για Excel)", value="Sheet1")

st.markdown("### 2) Γραμμές αρχείου")
c1, c2 = st.columns(2)
with c1:
    header_row = st.number_input("Header row (1-based)", value=1, min_value=1, step=1)
with c2:
    data_start_row = st.number_input("Data start row (1-based)", value=2, min_value=1, step=1)

st.markdown("### 3) STORE & BEX")
c3, c4 = st.columns(2)
with c3:
    store_letter = st.text_input("Γράμμα στήλης για STORE (άστο κενό για header aliases)", value="")
with c4:
    bex_mode = st.radio("BEX πηγή", ["Από λίστα stores", "Από στήλη (YES/NO)"], index=0)

manual_bex_list = st.text_input("Λίστα BEX stores (comma-separated)",
                                "DRZ01, FKM01, ESC01, LND01, PKK01")
bex_yesno_letter = ""
if bex_mode == "Από στήλη (YES/NO)":
    bex_yesno_letter = st.text_input("Γράμμα στήλης BEX (YES/NO)", value="", placeholder="π.χ. J")

st.markdown("### 4) Mapping με γράμματα")
cols = st.columns(4)
with cols[0]:
    letter_plan_vs      = st.text_input("plan_vs_target", value="A")
    letter_mobile_act   = st.text_input("mobile_actual", value="N")
    letter_llu          = st.text_input("llu_actual", value="T")
    letter_eon          = st.text_input("eon_tv_actual", value="X")
with cols[1]:
    letter_mobile_tgt   = st.text_input("mobile_target", value="O")
    letter_fixed_tgt    = st.text_input("fixed_target", value="P")
    letter_nga          = st.text_input("nga_actual", value="U")
    letter_fwa          = st.text_input("fwa_actual", value="Y")
with cols[2]:
    letter_fixed_act    = st.text_input("fixed_actual", value="Q")
    letter_voice_vs     = st.text_input("voice_vs_target", value="R")
    letter_ftth         = st.text_input("ftth_actual", value="V")
    letter_mob_upg      = st.text_input("mobile_upgrades", value="AA")
with cols[3]:
    letter_fixed_vs     = st.text_input("fixed_vs_target", value="S")
    letter_pending_mob  = st.text_input("pending_mobile", value="AF")
    letter_fixed_upg    = st.text_input("fixed_upgrades", value="AB")
    letter_pending_fix  = st.text_input("pending_fixed", value="AH")

run = st.button("🔧 Generate")

# ---------------- Run ----------------
if run:
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

    # Διαβάζουμε DataFrame ΜΟΝΟ για headers (και πιθανό STORE μέσω aliases)
    df = None
    wb = None
    ws = None
    try:
        fname = getattr(xls, "name", "")
        if fname.lower().endswith(".csv"):
            # CSV: δεν έχει sheets, δεν έχει openpyxl, τα γράμματα δεν έχουν έννοια → διαβάζουμε μόνο με pandas
            df = pd.read_csv(xls, header=header_row-1)
            st.write("📑 CSV headers:", list(df.columns))
        else:
            # XLSX: και pandas (για headers), και openpyxl (για γράμματα)
            xls_bytes = xls.read()
            xls_buf = io.BytesIO(xls_bytes)

            wb = load_workbook(xls_buf, data_only=True, read_only=True)
            if sheet_name not in wb.sheetnames:
                st.error(f"Το sheet '{sheet_name}' δεν βρέθηκε. Διαθέσιμα: {wb.sheetnames}")
                st.stop()
            ws = wb[sheet_name]

            # Για DataFrame headers: ξανα-ανοίγουμε δεύτερο buffer για pandas
            df = pd.read_excel(io.BytesIO(xls_bytes), sheet_name=sheet_name, engine="openpyxl", header=header_row-1)
            st.write("📑 XLSX headers:", list(df.columns))

    except Exception as e:
        st.error(f"Δεν άνοιξε το αρχείο: {e}")
        st.stop()

    if df is None or (isinstance(df, pd.DataFrame) and df.empty):
        st.error("Δεν βρέθηκαν δεδομένα.")
        st.stop()

    # Preview (2η γραμμή data = data_start_row)
    with st.expander("🔎 Preview (τραβάμε ΑΠΟ openpyxl με γράμματα)"):
        st.write("Headers (pandas):", list(df.columns))
        if ws is not None:
            preview_row = data_start_row  # 1-based excel row
            store_val_preview = ""
            if store_letter.strip():
                store_val_preview = get_value_by_letter(ws, store_letter, preview_row)
                store_header_preview = f"(by letter {store_letter})"
            else:
                # Δοκίμασε aliases στο pandas header
                aliases = ["Dealer_Code", "Dealer code", "Shop Code", "Shop_Code", "ShopCode", "Shop code",
                           "STORE", "Κατάστημα", "store", "dealer_code"]
                col = pick(df.columns, *aliases)
                if col:
                    store_val_preview = df.iloc[preview_row - data_start_row][col] if (preview_row - data_start_row) < len(df) else ""
                    store_header_preview = col
                else:
                    store_header_preview = "(no store header found)"

            prev = {
                "row_excel": preview_row,
                "store": {"from": store_header_preview, "value": str(store_val_preview)},
                "plan_vs_target": get_value_by_letter(ws, letter_plan_vs, preview_row),
                "mobile_actual":  get_value_by_letter(ws, letter_mobile_act, preview_row),
                "mobile_target":  get_value_by_letter(ws, letter_mobile_tgt, preview_row),
                "fixed_target":   get_value_by_letter(ws, letter_fixed_tgt, preview_row),
                "fixed_actual":   get_value_by_letter(ws, letter_fixed_act, preview_row),
                "voice_vs_target":get_value_by_letter(ws, letter_voice_vs, preview_row),
                "fixed_vs_target":get_value_by_letter(ws, letter_fixed_vs, preview_row),
                "llu_actual":     get_value_by_letter(ws, letter_llu, preview_row),
                "nga_actual":     get_value_by_letter(ws, letter_nga, preview_row),
                "ftth_actual":    get_value_by_letter(ws, letter_ftth, preview_row),
                "eon_tv_actual":  get_value_by_letter(ws, letter_eon, preview_row),
                "fwa_actual":     get_value_by_letter(ws, letter_fwa, preview_row),
                "mobile_upgrades":get_value_by_letter(ws, letter_mob_upg, preview_row),
                "fixed_upgrades": get_value_by_letter(ws, letter_fixed_upg, preview_row),
                "pending_mobile": get_value_by_letter(ws, letter_pending_mob, preview_row),
                "pending_fixed":  get_value_by_letter(ws, letter_pending_fix, preview_row),
            }
            st.json(prev)

    # Templates
    tpl_bex_bytes    = tpl_bex.read()
    tpl_nonbex_bytes = tpl_nonbex.read()

    out_zip = io.BytesIO()
    zf = zipfile.ZipFile(out_zip, "w", zipfile.ZIP_DEFLATED)
    built = 0

    # BEX set
    bex_set = set(s.strip().upper() for s in manual_bex_list.split(",") if s.strip())

    # Πόσες γραμμές να τρέξουμε
    total_rows = len(df) if not test_mode else min(50, len(df))

    # Λήψη STORE για γραμμή i (0-based πάνω στο df, αλλά openpyxl θέλει 1-based)
    def get_store_for_row(i_zero_based: int) -> str:
        row_excel_1based = data_start_row + i_zero_based
        if store_letter.strip() and ws is not None:
            v = get_value_by_letter(ws, store_letter, row_excel_1based)
            return (v or "").strip().upper()
        # αλλιώς από header aliases (pandas)
        aliases = ["Dealer_Code", "Dealer code", "Shop Code", "Shop_Code", "ShopCode", "Shop code",
                   "STORE", "Κατάστημα", "store", "dealer_code"]
        col = pick(df.columns, *aliases)
        if not col:
            return ""
        v = df.iloc[i_zero_based][col]
        return "" if pd.isna(v) else str(v).strip().upper()

    def val(letter: Optional[str], i_zero_based: int) -> str:
        if ws is None:
            # CSV path: προσπαθώ με pandas using column letters ≠ διαθέσιμο → επιστρέφω κενό
            return ""
        row_excel_1based = data_start_row + i_zero_based
        return get_value_by_letter(ws, letter, row_excel_1based)

    pbar = st.progress(0, text="Δημιουργία εγγράφων…")
    for i in range(total_rows):
        try:
            store_up = get_store_for_row(i)
            if not store_up:
                pbar.progress(min((i+1)/total_rows, 1.0), text=f"Παράλειψη γραμμής {i+1} (κενό store)")
                continue

            # BEX flag
            if bex_mode == "Από λίστα stores":
                is_bex = store_up in bex_set
                bex_text = "YES" if is_bex else "NO"
            else:
                # από στήλη YES/NO με γράμμα
                raw = val(bex_yesno_letter, i).strip().lower()
                is_bex = raw in ("yes", "y", "1", "true", "ναι")
                bex_text = "YES" if is_bex else "NO"

            mapping = {
                "title": f"Review September 2025 — Plan October 2025 — {store_up}",
                "plan_month": "Review September 2025 — Plan October 2025",
                "store": store_up,
                "bex": bex_text,

                "plan_vs_target":   safe(val(letter_plan_vs, i)),
                "mobile_actual":    safe(val(letter_mobile_act, i)),
                "mobile_target":    safe(val(letter_mobile_tgt, i)),
                "fixed_target":     safe(val(letter_fixed_tgt, i)),
                "fixed_actual":     safe(val(letter_fixed_act, i)),
                "voice_vs_target":  safe(val(letter_voice_vs, i)),
                "fixed_vs_target":  safe(val(letter_fixed_vs, i)),
                "llu_actual":       safe(val(letter_llu, i)),
                "nga_actual":       safe(val(letter_nga, i)),
                "ftth_actual":      safe(val(letter_ftth, i)),
                "eon_tv_actual":    safe(val(letter_eon, i)),
                "fwa_actual":       safe(val(letter_fwa, i)),
                "mobile_upgrades":  safe(val(letter_mob_upg, i)),
                "fixed_upgrades":   safe(val(letter_fixed_upg, i)),
                "pending_mobile":   safe(val(letter_pending_mob, i)),
                "pending_fixed":    safe(val(letter_pending_fix, i)),
            }

            doc = Document(io.BytesIO(tpl_bex_bytes if is_bex else tpl_nonbex_bytes))
            set_default_font(doc, "Aptos")
            replace_placeholders(doc, mapping)

            out_name = f"{store_up}_ReviewSep_PlanOct.docx"
            buf = io.BytesIO()
            doc.save(buf)
            zf.writestr(out_name, buf.getvalue())
            built += 1
            pbar.progress(min((i+1)/total_rows, 1.0), text=f"Φτιάχνω: {out_name} ({min(i+1, total_rows)}/{total_rows})")

        except Exception as e:
            st.warning(f"⚠️ Σφάλμα στη γραμμή {i+1}: {e}")
            if debug_mode:
                st.exception(e)

    zf.close()
    pbar.empty()

    if built == 0:
        st.error("Δεν δημιουργήθηκε αρχείο. Έλεγξε STORE/Data rows/Letters & templates.")
    else:
        st.success(f"Έτοιμα {built} αρχεία.")
        st.download_button("⬇️ Κατέβασε ZIP", data=out_zip.getvalue(), file_name="reviews_from_excel.zip")