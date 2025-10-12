# app.py
# Streamlit: Excel/CSV -> (BEX / Non-BEX) Review-Plan .docx (ZIP)

import io
import re
import zipfile
from typing import Any, Dict, Optional

import pandas as pd
import streamlit as st
from docx import Document
from docx.oxml.ns import qn

# ───────────── UI setup ─────────────
st.set_page_config(page_title="Excel → Review/Plan Generator", layout="wide")
st.title("📊 Excel/CSV → 📄 Review/Plan Generator (BEX & Non-BEX)")

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
    pattern = re.compile(r"\[\[([A-Za-z0-9_]+)\]\]")

    def sub_text(s: str) -> str:
        return pattern.sub(lambda m: "" if mapping.get(m.group(1)) is None else str(mapping.get(m.group(1), "")), s)

    for p in doc.paragraphs:
        for r in p.runs:
            r.text = sub_text(r.text)
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        r.text = sub_text(r.text)

def normkey(x: str) -> str:
    return re.sub(r"[\s\-_\.]+", "", str(x).strip().lower())

def pick_header(columns, *aliases) -> Optional[str]:
    nmap = {normkey(c): c for c in columns}
    for a in aliases:
        if normkey(a) in nmap:
            return nmap[normkey(a)]
    for a in aliases:
        pat = re.compile(a, re.IGNORECASE)
        for c in columns:
            if re.search(pat, str(c or "")):
                return c
    return None

def xl_letter_to_idx(letter: str) -> Optional[int]:
    """Excel col letters -> zero-based index. 'A'->0, 'Z'->25, 'AA'->26, 'AB'->27, etc."""
    s = (letter or "").strip().upper()
    if not s or not re.fullmatch(r"[A-Z]+", s):
        return None
    n = 0
    for ch in s:
        n = n * 26 + (ord(ch) - 64)
    return n - 1

def val_by_letter(row: pd.Series, df: pd.DataFrame, letter: Optional[str]) -> Any:
    """Δώσε γράμμα στήλης Excel, πάρε τιμή από τη σειρά (ασφαλές για ΑΑ, ΑΒ…)."""
    if not letter:
        return ""
    idx = xl_letter_to_idx(letter)
    if idx is None:
        return ""
    if 0 <= idx < len(df.columns):
        v = row.iloc[idx]
        return "" if pd.isna(v) else v
    return ""

def read_data(upload, sheet_name: str) -> Optional[pd.DataFrame]:
    try:
        fname = getattr(upload, "name", "")
        if fname.lower().endswith(".csv"):
            st.write("📑 Sheets:", ["CSV Data"])
            return pd.read_csv(upload)
        # xlsx
        xfile = pd.ExcelFile(upload, engine="openpyxl")
        st.write("📑 Sheets:", xfile.sheet_names)
        if sheet_name not in xfile.sheet_names:
            st.error(f"Το sheet '{sheet_name}' δεν βρέθηκε. Διαθέσιμα: {xfile.sheet_names}")
            return None
        return pd.read_excel(xfile, sheet_name=sheet_name, engine="openpyxl")
    except Exception as e:
        st.error(f"Δεν άνοιξε το αρχείο: {e}")
        return None

# ───────────── sidebar ─────────────
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

# ───────────── inputs ─────────────
st.markdown("### 1) Ανέβασε Excel/CSV")
xls = st.file_uploader("Excel/CSV", type=["xlsx", "csv"])
sheet_name = st.text_input("Όνομα φύλλου (για Excel)", value="Sheet1")

with st.expander("📌 Manual mapping με γράμματα Excel"):
    st.caption("Γράψε γράμματα στηλών (π.χ. A, N, AA, AB, AF, AH).")
    # Store (αν έχεις στήλη Dealer_Code ή Shop Code μπορείς να αφήσεις κενό και θα βρεθεί αυτόματα)
    store_letter = st.text_input("STORE letter (προαιρετικό, π.χ. A)", value="")

    # Τα γράμματα από το μήνυμα σου:
    plan_vs_target   = st.text_input("plan_vs_target", value="A")
    mobile_target    = st.text_input("mobile_target", value="O")
    fixed_target     = st.text_input("fixed_target",  value="P")
    fixed_actual     = st.text_input("fixed_actual",  value="Q")
    voice_vs_target  = st.text_input("voice_vs_target", value="R")
    fixed_vs_target  = st.text_input("fixed_vs_target", value="S")
    mobile_actual    = st.text_input("mobile_actual",  value="N")
    llu_actual       = st.text_input("llu_actual",     value="T")
    nga_actual       = st.text_input("nga_actual",     value="U")
    ftth_actual      = st.text_input("ftth_actual",    value="V")
    eon_tv_actual    = st.text_input("eon_tv_actual",  value="X")
    fwa_actual       = st.text_input("fwa_actual",     value="Y")
    mobile_upgrades  = st.text_input("mobile_upgrades", value="AA")
    fixed_upgrades   = st.text_input("fixed_upgrades",  value="AB")
    pending_mobile   = st.text_input("pending_mobile",  value="AF")
    pending_fixed    = st.text_input("pending_fixed",   value="AH")

with st.expander("🏷️ BEX detection"):
    bex_mode = st.radio("Πώς βρίσκουμε αν είναι BEX;", ["Λίστα καταστημάτων", "Από στήλη (YES/NO)"], index=0, horizontal=True)
    bex_list_input = st.text_input(
        "Σταθερή λίστα (comma-separated)",
        value="DRZ01,FKM01,ESC01,LND01,PKK01"
    )
    bex_yesno_letter = st.text_input("Letter στήλης BEX (YES/NO), αν επέλεξες 'Από στήλη'", value="")

# ───────────── main ─────────────
run = st.button("🔧 Generate")

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

    df = read_data(xls, sheet_name)
    if df is None or df.empty:
        st.error("Δεν βρέθηκαν δεδομένα.")
        st.stop()

    st.success(f"OK: {len(df)} γραμμές, {len(df.columns)} στήλες.")
    if debug_mode:
        st.write("Headers:", list(df.columns))
        st.dataframe(df.head(10))

    # Βρες STORE header αν δεν δόθηκε γράμμα
    store_hdr = None
    if not store_letter.strip():
        store_hdr = pick_header(
            df.columns,
            "Dealer_Code", "Dealer code", "Shop Code", "Shop_Code", "ShopCode",
            "STORE", "Κατάστημα", r"shop.?code", r"dealer.?code"
        )

    # Preview mapping από 2η γραμμή (αν υπάρχει)
    with st.expander("🔍 Mapping preview (από 2η γραμμή)"):
        preview_obj = {}
        row2 = df.iloc[1] if len(df) > 1 else df.iloc[0]
        def sample(letter):
            return val_by_letter(row2, df, letter)
        preview_obj.update({
            "store_letter": {
                "letter": store_letter or "(auto header)",
                "header": store_hdr,
                "sample_row2": (row2[store_hdr] if store_hdr else sample(store_letter)) if len(df) else ""
            },
            "plan_vs_target": {"letter": plan_vs_target, "sample_row2": sample(plan_vs_target)},
            "mobile_actual":  {"letter": mobile_actual,  "sample_row2": sample(mobile_actual)},
            "mobile_target":  {"letter": mobile_target,  "sample_row2": sample(mobile_target)},
            "fixed_target":   {"letter": fixed_target,   "sample_row2": sample(fixed_target)},
            "fixed_actual":   {"letter": fixed_actual,   "sample_row2": sample(fixed_actual)},
            "voice_vs_target":{"letter": voice_vs_target,"sample_row2": sample(voice_vs_target)},
            "fixed_vs_target":{"letter": fixed_vs_target,"sample_row2": sample(fixed_vs_target)},
            "llu_actual":     {"letter": llu_actual,     "sample_row2": sample(llu_actual)},
            "nga_actual":     {"letter": nga_actual,     "sample_row2": sample(nga_actual)},
            "ftth_actual":    {"letter": ftth_actual,    "sample_row2": sample(ftth_actual)},
            "eon_tv_actual":  {"letter": eon_tv_actual,  "sample_row2": sample(eon_tv_actual)},
            "fwa_actual":     {"letter": fwa_actual,     "sample_row2": sample(fwa_actual)},
            "mobile_upgrades":{"letter": mobile_upgrades,"sample_row2": sample(mobile_upgrades)},
            "fixed_upgrades": {"letter": fixed_upgrades, "sample_row2": sample(fixed_upgrades)},
            "pending_mobile": {"letter": pending_mobile, "sample_row2": sample(pending_mobile)},
            "pending_fixed":  {"letter": pending_fixed,  "sample_row2": sample(pending_fixed)},
        })
        st.write(preview_obj)

    # Templates σε bytes
    tpl_bex_bytes    = tpl_bex.read()
    tpl_nonbex_bytes = tpl_nonbex.read()

    # ZIP out
    out_zip = io.BytesIO()
    zf = zipfile.ZipFile(out_zip, "w", zipfile.ZIP_DEFLATED)
    built = 0

    pbar = st.progress(0, text="Δημιουργία εγγράφων…")
    total = len(df) if not test_mode else min(50, len(df))

    # BEX list set
    bex_set = set(s.strip().upper() for s in bex_list_input.split(",") if s.strip())

    for i in range(len(df)):
        if test_mode and i >= total:
            st.info(f"🧪 Test mode: σταμάτησα στις {total} γραμμές.")
            break

        row = df.iloc[i]

        # STORE
        if store_letter.strip():
            store = str(val_by_letter(row, df, store_letter)).strip()
        else:
            store = str(row.get(store_hdr, "")).strip() if store_hdr else ""
        if not store:
            pbar.progress(min((i + 1) / (total or 1), 1.0), text=f"Παράλειψη γραμμής {i+1} (κενό store)")
            continue
        store_up = store.upper()

        # BEX?
        if bex_mode == "Λίστα καταστημάτων":
            is_bex = store_up in bex_set
            bex_str = "YES" if is_bex else "NO"
        else:
            bex_val = str(val_by_letter(row, df, bex_yesno_letter)).strip().lower()
            is_bex = bex_val in ("yes", "y", "1", "true", "ναι")
            bex_str = "YES" if is_bex else "NO"

        # Fields
        def g(letter): 
            v = val_by_letter(row, df, letter)
            return "" if pd.isna(v) else v

        mapping = {
            "title": f"Review September 2025 — Plan October 2025 — {store_up}",
            "plan_month": "Review September 2025 → Plan October 2025",
            "store": store_up,
            "bex": bex_str,

            "plan_vs_target": g(plan_vs_target),
            "mobile_actual":  g(mobile_actual),
            "mobile_target":  g(mobile_target),
            "fixed_target":   g(fixed_target),
            "fixed_actual":   g(fixed_actual),
            "voice_vs_target":g(voice_vs_target),
            "fixed_vs_target":g(fixed_vs_target),
            "llu_actual":     g(llu_actual),
            "nga_actual":     g(nga_actual),
            "ftth_actual":    g(ftth_actual),
            "eon_tv_actual":  g(eon_tv_actual),
            "fwa_actual":     g(fwa_actual),
            "mobile_upgrades":g(mobile_upgrades),
            "fixed_upgrades": g(fixed_upgrades),
            "pending_mobile": g(pending_mobile),
            "pending_fixed":  g(pending_fixed),
        }

        try:
            doc = Document(io.BytesIO(tpl_bex_bytes if is_bex else tpl_nonbex_bytes))
            set_default_font(doc, "Aptos")
            replace_placeholders(doc, mapping)

            out_name = f"{store_up}_ReviewSep_PlanOct.docx"
            buf = io.BytesIO()
            doc.save(buf)
            zf.writestr(out_name, buf.getvalue())
            built += 1

            pbar.progress(min((i + 1) / (total or 1), 1.0), text=f"Φτιάχνω: {out_name} ({min(i+1, total)}/{total})")

        except Exception as e:
            st.warning(f"⚠️ Σφάλμα στη γραμμή {i+1}: {e}")
            if debug_mode:
                st.exception(e)

    zf.close()
    pbar.empty()

    if built == 0:
        st.error("Δεν δημιουργήθηκε αρχείο. Έλεγξε STORE/BEX mapping & templates.")
    else:
        st.success(f"Έτοιμα {built} αρχεία.")
        st.download_button("⬇️ Κατέβασε ZIP", data=out_zip.getvalue(), file_name="reviews_from_excel.zip")