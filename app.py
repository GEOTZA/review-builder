# app.py — Streamlit: Excel/CSV -> (BEX / Non-BEX) Review-Plan .docx (ZIP)

import io
import re
import zipfile
import json
import unicodedata
from typing import Any, Dict, Optional, List

import pandas as pd
import streamlit as st
from docx import Document
from docx.oxml.ns import qn


# ───────────────────────────── UI CONFIG ─────────────────────────────
st.set_page_config(page_title="Excel → Review/Plan Generator", layout="wide")
st.title("📊 Excel/CSV → 📄 Review/Plan Generator (BEX & Non-BEX)")


# ───────────────────────────── HELPERS ─────────────────────────────
def set_default_font(doc: Document, font_name: str = "Aptos") -> None:
    """Set default font for all styles (incl. eastAsia) to avoid font mismatches."""
    for style in doc.styles:
        if hasattr(style, "font"):
            try:
                style.font.name = font_name
                style._element.rPr.rFonts.set(qn("w:eastAsia"), font_name)
                style._element.rPr.rFonts.set(qn("w:cs"), font_name)
            except Exception:
                pass


def replace_placeholders(doc: Document, mapping: Dict[str, Any]) -> None:
    """Replace [[placeholders]] across paragraphs and tables."""
    pattern = re.compile(r"\[\[([A-Za-z0-9_]+)\]\]")

    def subfun(s: str) -> str:
        def val(m):
            key = m.group(1)
            v = mapping.get(key, "")
            return "" if v is None else str(v)
        return pattern.sub(val, s)

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
    """
    Aggressive header normalization:
    - NFKD unicode normalize (removes accents)
    - lowercase
    - remove ALL non [a-z0-9]
    """
    s = unicodedata.normalize("NFKD", str(x))
    s = s.lower()
    s = re.sub(r"[^a-z0-9]+", "", s)
    return s


def pick(columns: List[str], *aliases: str) -> Optional[str]:
    """Find a column by (normalized) aliases, then by regex contains."""
    nmap = {normkey(c): c for c in columns}
    # exact by normalized
    for a in aliases:
        if normkey(a) in nmap:
            return nmap[normkey(a)]
    # regex contains on the original headers
    for a in aliases:
        try:
            pat = re.compile(a, re.IGNORECASE)
        except re.error:
            continue
        for c in columns:
            if re.search(pat, str(c)):
                return c
    return None


def cell(row: pd.Series, col: Optional[str]):
    if not col:
        return ""
    v = row[col]
    return "" if pd.isna(v) else v


def read_data(file, sheet_name: str) -> Optional[pd.DataFrame]:
    """
    Accepts .xlsx or .csv (auto-detect by uploaded name).
    Returns DataFrame or None (and shows a user-friendly error).
    """
    try:
        fname = getattr(file, "name", "")
        if fname.lower().endswith(".csv"):
            st.write("📑 Sheets:", ["CSV Data"])
            return pd.read_csv(file)

        # default: xlsx
        xfile = pd.ExcelFile(file, engine="openpyxl")
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
test_mode = st.sidebar.toggle("🧪 Test mode (limit rows=50)", value=True)

st.sidebar.header("⚙️ BEX")
bex_mode = st.sidebar.radio("Πηγή BEX", ["Στήλη στο Excel", "Λίστα (comma-separated)"], index=0)
bex_list = set()
if bex_mode == "Λίστα (comma-separated)":
    bex_input = st.sidebar.text_area("BEX stores", "ESC01,FKM01,LND01,DRZ01,PKK01")
    bex_list = set(s.strip().upper() for s in bex_input.split(",") if s.strip())

st.sidebar.subheader("📄 Templates (.docx)")
tpl_bex = st.sidebar.file_uploader("BEX template", type=["docx"])
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
    # Basic checks
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

    # Read data
    df = read_data(xls, sheet_name)
    if df is None or df.empty:
        st.error("Δεν βρέθηκαν δεδομένα στο αρχείο.")
        st.stop()

    st.success(f"OK: {len(df)} γραμμές, {len(df.columns)} στήλες.")

    # ---- Εμφάνιση & Κατέβασμα κεφαλίδων ----
    headers = list(df.columns)
    if debug_mode:
        st.markdown("#### 📋 Headers όπως τους βλέπουμε:")
        st.code(headers)
        colh1, colh2 = st.columns(2)
        with colh1:
            st.download_button("⬇️ Κατέβασε Headers (.txt)",
                               "\n".join(map(str, headers)).encode("utf-8"),
                               file_name="headers.txt")
        with colh2:
            st.download_button("⬇️ Κατέβασε Headers (.json)",
                               json.dumps(list(map(str, headers)), ensure_ascii=False, indent=2).encode("utf-8"),
                               file_name="headers.json")
        st.dataframe(df.head(10))

    cols = list(map(str, df.columns))

    # ── Auto-map
    col_store = pick(
        cols,
        # Αγγλικά
        "Shop Code", "Shop_Code", "ShopCode", "Shop code",
        "Store Code", "Store_Code", "StoreCode",
        "Dealer Code", "Dealer_Code", "DealerCode",
        # Ελληνικά
        "Κωδικός Καταστήματος", "Κωδικος Καταστηματος", "Κωδικός", "Κατάστημα", "Καταστημα",
        # Regex
        r"shop.?code", r"store.?code", r"dealer.?code"
    )
    col_bex = pick(cols, "BEX store", "BEX", r"bex.?store")
    col_mob_act = pick(cols, "mobile actual", r"mobile.*actual")
    col_mob_tgt = pick(cols, "mobile target", r"mobile.*target", "mobile plan")
    col_fix_tgt = pick(cols, "target fixed", r"fixed.*target", "fixed plan total", "fixed plan")
    col_fix_act = pick(cols, "total fixed", r"(total|sum).?fixed.*actual", "fixed actual")
    col_pend_mob = pick(cols, "TOTAL PENDING MOBILE", r"pending.*mobile")
    col_pend_fix = pick(cols, "TOTAL PENDING FIXED", r"pending.*fixed")
    col_plan_vs = pick(cols, "plan vs target", r"plan.*vs.*target")

    with st.expander("Χαρτογράφηση (auto)"):
        st.write({
            "STORE": col_store, "BEX": col_bex,
            "mobile_actual": col_mob_act, "mobile_target": col_mob_tgt,
            "fixed_target": col_fix_tgt, "fixed_actual": col_fix_act,
            "pending_mobile": col_pend_mob, "pending_fixed": col_pend_fix,
            "plan_vs_target": col_plan_vs
        })

    # ---- Manual mapping (WITH FORM + SESSION STATE)
    missing = []
    if not col_store: missing.append("STORE")
    if not col_mob_act: missing.append("mobile_actual")
    if not col_mob_tgt: missing.append("mobile_target")
    if not col_fix_act: missing.append("fixed_actual")
    if not col_fix_tgt: missing.append("fixed_target")
    if not col_pend_mob: missing.append("pending_mobile")
    if not col_pend_fix: missing.append("pending_fixed")
    if not col_plan_vs: missing.append("plan_vs_target")

    def _init_key(k, v):
        if k not in st.session_state:
            st.session_state[k] = v or ""

    _init_key("map_STORE", col_store or "")
    _init_key("map_BEX", col_bex or "")
    _init_key("map_mobile_actual", col_mob_act or "")
    _init_key("map_mobile_target", col_mob_tgt or "")
    _init_key("map_fixed_actual", col_fix_act or "")
    _init_key("map_fixed_target", col_fix_tgt or "")
    _init_key("map_pending_mobile", col_pend_mob or "")
    _init_key("map_pending_fixed", col_pend_fix or "")
    _init_key("map_plan_vs_target", col_plan_vs or "")
    _init_key("mapping_locked", False)

    if missing or not st.session_state["mapping_locked"]:
        st.info("Ρύθμισε/κλείδωσε τη χαρτογράφηση (δεν θα ξανατρέχει σε κάθε αλλαγή).")
        options = [""] + [str(c) for c in cols]

        with st.form("mapping_form"):
            c1, c2, c3 = st.columns(3)
            with c1:
                st.selectbox("STORE (Shop/Dealer code)", options, key="map_STORE")
                st.selectbox("BEX flag (Yes/No)", options, key="map_BEX")
                st.selectbox("plan_vs_target (%)", options, key="map_plan_vs_target")
            with c2:
                st.selectbox("mobile_actual", options, key="map_mobile_actual")
                st.selectbox("mobile_target", options, key="map_mobile_target")
                st.selectbox("pending_mobile", options, key="map_pending_mobile")
            with c3:
                st.selectbox("fixed_actual", options, key="map_fixed_actual")
                st.selectbox("fixed_target", options, key="map_fixed_target")
                st.selectbox("pending_fixed", options, key="map_pending_fixed")

            submitted = st.form_submit_button("✅ Χρήση χαρτογράφησης")
            if submitted:
                st.session_state["mapping_locked"] = True
                st.success("Η χαρτογράφηση κλειδώθηκε. Προχώρα στην παραγωγή αρχείων.")

    # τελικά χρησιμοποιούμε ό,τι είναι στο session_state (ή το auto-map αν έμεινε κενό)
    col_store     = st.session_state["map_STORE"] or col_store
    col_bex       = st.session_state["map_BEX"] or col_bex
    col_plan_vs   = st.session_state["map_plan_vs_target"] or col_plan_vs
    col_mob_act   = st.session_state["map_mobile_actual"] or col_mob_act
    col_mob_tgt   = st.session_state["map_mobile_target"] or col_mob_tgt
    col_pend_mob  = st.session_state["map_pending_mobile"] or col_pend_mob
    col_fix_act   = st.session_state["map_fixed_actual"] or col_fix_act
    col_fix_tgt   = st.session_state["map_fixed_target"] or col_fix_tgt
    col_pend_fix  = st.session_state["map_pending_fixed"] or col_pend_fix

    if not col_store:
        st.error("Διάλεξε στήλη STORE και πάτα ‘Χρήση χαρτογράφησης’.")
        st.stop()

    # ── Templates
    tpl_bex_bytes = tpl_bex.read()
    tpl_nonbex_bytes = tpl_nonbex.read()

    # ── Build ZIP
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
                # skip row if store empty
                if (i % 10 == 0) or (i == total):
                    pbar.progress(min(i / (total or 1), 1.0), text=f"Παράλειψη γραμμής {i} (κενό store)")
                continue

            store_up = store.upper()

            # BEX flag
            if bex_mode == "Λίστα (comma-separated)":
                is_bex = store_up in bex_list
            else:
                bex_val = str(cell(row, col_bex)).strip().lower() if col_bex else "no"
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

            # build docx
            doc = Document(io.BytesIO(tpl_bex_bytes if is_bex else tpl_nonbex_bytes))
            set_default_font(doc, "Aptos")
            replace_placeholders(doc, mapping)

            out_name = f"{store_up}_ReviewSep_PlanOct.docx"
            buf = io.BytesIO()
            doc.save(buf)
            zf.writestr(out_name, buf.getvalue())
            built += 1

            if (i % 10 == 0) or (i == total):
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
        st.success(f"Έτοιμα {built} αρχεία.")
        st.download_button("⬇️ Κατέβασε ZIP", data=out_zip.getvalue(), file_name="reviews_from_excel.zip")