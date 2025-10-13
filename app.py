λ# app.py
# Streamlit: Excel -> (BEX / Non-BEX) Review-Plan .docx (ZIP)

import io
import re
import zipfile
from typing import Any, Dict, Optional

import streamlit as st
from docx import Document
from docx.oxml.ns import qn
from openpyxl import load_workbook

# ───────────── UI setup ─────────────
st.set_page_config(page_title="Excel → Review/Plan Generator", layout="wide")
st.title("📊 Excel → 📄 Review/Plan Generator (BEX & Non-BEX)")

# ───────────── helpers ─────────────
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
    pat = re.compile(r"\[\[([A-Za-z0-9_]+)\]\]")
    def sub(s: str) -> str:
        return pat.sub(lambda m: "" if mapping.get(m.group(1)) is None else str(mapping.get(m.group(1), "")), s)

    for p in doc.paragraphs:
        for r in p.runs:
            r.text = sub(r.text)
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        r.text = sub(r.text)

def normkey(x: Any) -> str:
    return re.sub(r"[\s\-_\.]+", "", str(x).strip().lower())

def col_letter_to_index(letter: str) -> Optional[int]:
    """Excel letters -> 1-based index: A->1, Z->26, AA->27, ..."""
    if not letter:
        return None
    s = str(letter).strip().upper()
    if not re.fullmatch(r"[A-Z]+", s):
        return None
    n = 0
    for ch in s:
        n = n*26 + (ord(ch) - 64)
    return n  # 1-based

def read_cell(ws, row_idx: int, col_letter: Optional[str]):
    if not col_letter:
        return ""
    c = col_letter_to_index(col_letter)
    if not c:
        return ""
    v = ws.cell(row=row_idx, column=c).value
    return "" if v is None else v

def find_header_col_by_names(ws, header_row: int, candidates: list[str]) -> Optional[int]:
    """Βρες column index (1-based) στη γραμμή header_row που ταιριάζει με κάποια από τις ονομασίες."""
    max_col = ws.max_column
    cand_norm = [normkey(x) for x in candidates]
    for col in range(1, max_col+1):
        hdr = ws.cell(row=header_row, column=col).value
        if hdr is None:
            continue
        h = normkey(hdr)
        if any(a == h for a in cand_norm):
            return col
        # contains match για safety
        if any(a and a in h for a in cand_norm):
            return col
    return None

def val_by_header(ws, row_idx: int, header_row: int, header_names: list[str]):
    """Δώσε row + λίστα πιθανών headers, πάρε τιμή (αν βρεθεί header)."""
    col = find_header_col_by_names(ws, header_row, header_names)
    if not col:
        return ""
    v = ws.cell(row=row_idx, column=col).value
    return "" if v is None else v

# ───────────── sidebar controls ─────────────
with st.sidebar:
    st.subheader("🛠 Ρυθμίσεις")
    debug_mode = st.toggle("Debug mode", value=True)
    test_mode  = st.toggle("Test mode (πρώτες 50 γραμμές)", value=True)

    st.subheader("📄 Templates (.docx)")
    tpl_bex    = st.file_uploader("BEX template", type=["docx"])
    tpl_nonbex = st.file_uploader("Non-BEX template", type=["docx"])
    st.caption(
        "Placeholders: [[title]], [[plan_month]], [[store]], [[bex]], "
        "[[plan_vs_target]], [[mobile_actual]], [[mobile_target]], "
        "[[fixed_actual]], [[fixed_target]], [[voice_vs_target]], [[fixed_vs_target]], "
        "[[llu_actual]], [[nga_actual]], [[ftth_actual]], [[eon_tv_actual]], [[fwa_actual]], "
        "[[mobile_upgrades]], [[fixed_upgrades]], [[pending_mobile]], [[pending_fixed]]"
    )

# ───────────── main inputs ─────────────
st.markdown("### 1) Ανέβασε Excel")
xls = st.file_uploader("Excel (.xlsx)", type=["xlsx"])
sheet_name = st.text_input("Όνομα φύλλου (Sheet)", value="Sheet1")

with st.expander("📌 Ρύθμιση γραμμών (headers & δεδομένα)"):
    header_row = st.number_input("Header row (1-based)", min_value=1, value=1, step=1)
    data_start_row = st.number_input("First data row (1-based)", min_value=1, value=2, step=1)

with st.expander("🏷️ STORE & BEX"):
    store_letter = st.text_input("STORE letter (προαιρετικό). Αν το αφήσεις κενό, θα ψάξω header.", value="")
    st.caption("Αν δεν βάλεις γράμμα, θα ψάξει headers όπως: Dealer_Code, Shop Code, Store, Κατάστημα στη γραμμή headers.")

    bex_mode = st.radio("BEX:", ["Από λίστα καταστημάτων", "Από στήλη (YES/NO)"], index=0, horizontal=True)
    bex_list_input = st.text_input("Σταθερή λίστα BEX (comma-separated)", value="DRZ01,FKM01,ESC01,LND01,PKK01")
    bex_yesno_letter = st.text_input("Letter στήλης BEX (YES/NO) — αν διάλεξες 'Από στήλη'", value="")

with st.expander("🗺️ Mapping με γράμματα Excel (A, N, AA, AB, AF, AH)"):
    # Default letters από εσένα
    plan_vs_target   = st.text_input("plan_vs_target", value="A")
    mobile_actual    = st.text_input("mobile_actual",  value="N")
    mobile_target    = st.text_input("mobile_target",  value="O")
    fixed_target     = st.text_input("fixed_target",   value="P")
    fixed_actual     = st.text_input("fixed_actual",   value="Q")
    voice_vs_target  = st.text_input("voice_vs_target", value="R")
    fixed_vs_target  = st.text_input("fixed_vs_target", value="S")
    llu_actual       = st.text_input("llu_actual",     value="T")
    nga_actual       = st.text_input("nga_actual",     value="U")
    ftth_actual      = st.text_input("ftth_actual",    value="V")
    eon_tv_actual    = st.text_input("eon_tv_actual",  value="X")
    fwa_actual       = st.text_input("fwa_actual",     value="Y")
    mobile_upgrades  = st.text_input("mobile_upgrades", value="AA")
    fixed_upgrades   = st.text_input("fixed_upgrades",  value="AB")
    pending_mobile   = st.text_input("pending_mobile",  value="AF")
    pending_fixed    = st.text_input("pending_fixed",   value="AH")

run = st.button("🔧 Generate")

# ───────────── main logic ─────────────
if run:
    if not xls:
        st.error("Ανέβασε Excel (.xlsx) πρώτα.")
        st.stop()
    if not tpl_bex or not tpl_nonbex:
        st.error("Ανέβασε και τα δύο templates (.docx).")
        st.stop()

    # Άνοιγμα workbook/sheet για άμεση ανάγνωση κελιών
    try:
        wb = load_workbook(filename=xls, data_only=True)
        if sheet_name not in wb.sheetnames:
            st.error(f"Το sheet '{sheet_name}' δεν βρέθηκε. Διαθέσιμα: {wb.sheetnames}")
            st.stop()
        ws = wb[sheet_name]
    except Exception as e:
        st.error(f"Αδυναμία ανοίγματος Excel: {e}")
        st.stop()

    # Προετοιμασία STORE ανάγνωση
    store_header_candidates = [
        "Dealer_Code", "Dealer code", "Shop Code", "Shop_Code", "ShopCode",
        "Store", "STORE", "Κατάστημα", "Κωδικός Καταστήματος", "ΚΩΔΙΚΟΣ ΚΑΤΑΣΤΗΜΑΤΟΣ",
        r"shop.?code", r"dealer.?code"
    ]
    store_col_index = None
    if not store_letter.strip():
        store_col_index = find_header_col_by_names(ws, header_row, store_header_candidates)

    # Preview (δεύτερη γραμμή δεδομένων)
    with st.expander("🔍 Preview (από την πρώτη γραμμή δεδομένων)"):
        r = data_start_row
        def sample(letter): return read_cell(ws, r, letter)
        if store_letter.strip():
            store_preview = read_cell(ws, r, store_letter)
        else:
            store_preview = ws.cell(row=r, column=store_col_index).value if store_col_index else ""
        st.write({
            "store": store_preview,
            "plan_vs_target": sample(plan_vs_target),
            "mobile_actual": sample(mobile_actual),
            "mobile_target": sample(mobile_target),
            "fixed_target": sample(fixed_target),
            "fixed_actual": sample(fixed_actual),
            "voice_vs_target": sample(voice_vs_target),
            "fixed_vs_target": sample(fixed_vs_target),
            "llu_actual": sample(llu_actual),
            "nga_actual": sample(nga_actual),
            "ftth_actual": sample(ftth_actual),
            "eon_tv_actual": sample(eon_tv_actual),
            "fwa_actual": sample(fwa_actual),
            "mobile_upgrades": sample(mobile_upgrades),
            "fixed_upgrades": sample(fixed_upgrades),
            "pending_mobile": sample(pending_mobile),
            "pending_fixed": sample(pending_fixed),
        })

    # Templates
    tpl_bex_bytes    = tpl_bex.read()
    tpl_nonbex_bytes = tpl_nonbex.read()

    # ZIP out
    out_zip = io.BytesIO()
    zf = zipfile.ZipFile(out_zip, "w", zipfile.ZIP_DEFLATED)
    built = 0

    pbar = st.progress(0, text="Δημιουργία εγγράφων…")

    # BEX set
    bex_set = set(s.strip().upper() for s in bex_list_input.split(",") if s.strip())

    # Πόσες γραμμές θα τρέξουμε
    max_row = ws.max_row
    last_row = max_row
    if test_mode:
        last_row = min(max_row, data_start_row - 1 + 50)

    # Συνάρτηση για γρήγορη ανάγνωση πεδίου από γράμμα
    def G(row_idx: int, letter: str):
        return read_cell(ws, row_idx, letter)

    # Loop σε γραμμές δεδομένων
    total_rows = max(0, last_row - data_start_row + 1)
    cur = 0
    for row_idx in range(data_start_row, last_row + 1):
        cur += 1

        # STORE
        if store_letter.strip():
            store = str(G(row_idx, store_letter)).strip()
        else:
            store = str(ws.cell(row=row_idx, column=store_col_index).value if store_col_index else "").strip()

        # Αν είναι κενό το store, σταμάτα (συνήθως τέλος πίνακα)
        if not store:
            pbar.progress(min(cur / (total_rows or 1), 1.0), text=f"Stop στη γραμμή {row_idx} (κενό STORE)")
            break

        store_up = store.upper()

        # BEX;
        if bex_mode == "Από λίστα καταστημάτων":
            is_bex = store_up in bex_set
        else:
            bex_val = str(G(row_idx, bex_yesno_letter)).strip().lower() if bex_yesno_letter else ""
            is_bex = bex_val in ("yes", "y", "1", "true", "ναι")
        bex_str = "YES" if is_bex else "NO"

        # Mapping values
        mapping = {
            "title": f"Review September 2025 — Plan October 2025 — {store_up}",
            "plan_month": "Review September 2025 — Plan October 2025",
            "store": store_up,
            "bex": bex_str,

            "plan_vs_target":   G(row_idx, plan_vs_target),
            "mobile_actual":    G(row_idx, mobile_actual),
            "mobile_target":    G(row_idx, mobile_target),
            "fixed_target":     G(row_idx, fixed_target),
            "fixed_actual":     G(row_idx, fixed_actual),
            "voice_vs_target":  G(row_idx, voice_vs_target),
            "fixed_vs_target":  G(row_idx, fixed_vs_target),
            "llu_actual":       G(row_idx, llu_actual),
            "nga_actual":       G(row_idx, nga_actual),
            "ftth_actual":      G(row_idx, ftth_actual),
            "eon_tv_actual":    G(row_idx, eon_tv_actual),
            "fwa_actual":       G(row_idx, fwa_actual),
            "mobile_upgrades":  G(row_idx, mobile_upgrades),
            "fixed_upgrades":   G(row_idx, fixed_upgrades),
            "pending_mobile":   G(row_idx, pending_mobile),
            "pending_fixed":    G(row_idx, pending_fixed),
        }

        try:
            tpl_bytes = tpl_bex_bytes if is_bex else tpl_nonbex_bytes
            doc = Document(io.BytesIO(tpl_bytes))
            set_default_font(doc, "Aptos")
            replace_placeholders(doc, mapping)

            out_name = f"{store_up}_ReviewSep_PlanOct.docx"
            buf = io.BytesIO()
            doc.save(buf)
            zf.writestr(out_name, buf.getvalue())
            built += 1

            pbar.progress(min(cur / (total_rows or 1), 1.0), text=f"Φτιάχνω: {out_name} ({cur}/{total_rows})")
        except Exception as e:
            st.warning(f"⚠️ Σφάλμα στη γραμμή {row_idx}: {e}")
            if debug_mode:
                st.exception(e)

    zf.close()
    pbar.empty()

    if built == 0:
        st.error("Δεν δημιουργήθηκε αρχείο. Έλεγξε STORE/BEX mapping & templates.")
    else:
        st.success(f"Έτοιμα {built} αρχεία.")
        st.download_button("⬇️ Κατέβασε ZIP", data=out_zip.getvalue(), file_name="reviews_from_excel.zip")