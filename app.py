# app.py — Streamlit: Excel/CSV -> (BEX / Non-BEX) Review-Plan .docx (ZIP)

import io
import re
import zipfile
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
                # some styles don't have rPr
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
    if debug_mode:
        st.markdown("#### Headers όπως τους βλέπουμε:")
        st.code(list(df.columns))
        st.dataframe(df.head(10))

    cols = list(df.columns)

    # ── Auto-map
    col_store = pick(
        cols,
        # Συνήθη αγγλικά
        "Shop Code", "Shop_Code", "ShopCode", "Shop code",
        "Store Code", "Store_Code", "StoreCode",
        "Dealer Code", "Dealer_Code", "DealerCode",
        # Ελληνικά πιθανά
        "Κωδικός Καταστήματος", "Κωδικος Καταστηματος", "Κωδικός", "Κατάστημα", "Καταστημα",
        # Regex/γενικά
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

    # ── Manual mapping if something is missing
    missing = []
    if not col_store: missing.append("STORE")
    if not col_mob_act: missing.append("mobile_actual")
    if not col_mob_tgt: missing.append("mobile_target")
    if not col_fix_act: missing.append("fixed_actual")
    if not col_fix_tgt: missing.append("fixed_target")
    if not col_pend_mob: missing.append("pending_mobile")
    if not col_pend_fix: missing.append("pending_fixed")
    if not col_plan_vs: missing.append("plan_vs_target")

    if missing:
        st.warning("Κάποια πεδία δεν αναγνωρίστηκαν αυτόματα. Διάλεξε χειροκίνητα:")
        options = [""] + [str(c) for c in cols]

        c1, c2, c3 = st.columns(3)
        with c1:
            col_store = st.selectbox("STORE (Shop/Dealer code)", options,
                                     index=options.index(col_store) if col_store in options else 0) or None
            col_bex = st.selectbox("BEX flag (Yes/No)", options,
                                   index=options.index(col_bex) if col_bex in options else 0) or None
            col_plan_vs = st.selectbox("plan_vs_target (%)", options,
                                       index=options.index(col_plan_vs) if col_plan_vs in options else 0) or None
        with c2:
            col_mob_act = st.selectbox("mobile_actual", options,
                                       index=options.index(col_mob_act) if col_mob_act in options else 0) or None
            col_mob_tgt = st.selectbox("mobile_target", options,
                                       index=options.index(col_mob_tgt) if col_mob_tgt in options else 0) or None
            col_pend_mob = st.selectbox("pending_mobile", options,
                                        index=options.index(col_pend_mob) if col_pend_mob in options else 0) or None
        with c3:
            col_fix_act = st.selectbox("fixed_actual", options,
                                       index=options.index(col_fix_act) if col_fix_act in options else 0) or None
            col_fix_tgt = st.selectbox("fixed_target", options,
                                       index=options.index(col_fix_tgt) if col_fix_tgt in options else 0) or None
            col_pend_fix = st.selectbox("pending_fixed", options,
                                        index=options.index(col_pend_fix) if col_pend_fix in options else 0) or None

    if not col_store:
        st.error("Δεν βρέθηκε στήλη STORE (π.χ. 'Shop Code' ή 'Dealer_Code'). Διόρθωσε/διάλεξε από τα dropdowns.")
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

    # iterate rows
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
        st.success(f"Έτοιμα {built} αρχεία.")
        st.download_button("⬇️ Κατέβασε ZIP", data=out_zip.getvalue(), file_name="reviews_from_excel.zip")