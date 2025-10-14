# app.py
# Streamlit app: Excel → (BEX / NON-BEX) DOCX generator with robust placeholder replacement
# by you + helper ♥

import io
import re
import zipfile
import datetime as dt
from pathlib import Path
from typing import Any, Dict, Iterable

import streamlit as st
import pandas as pd
from docx import Document

# ───────────────────────────── Config ─────────────────────────────
st.set_page_config(page_title="Excel → Review/Plan (BEX & Non-BEX)", layout="wide")
TODAY = dt.date.today()
HERE = Path(__file__).parent

# ───────────────────────── Helpers ─────────────────────────
_RX_PH = re.compile(r"\[\[([A-Za-z0-9_]+)\]\]")  # [[key]]

def format_percent(val: Any) -> str:
    """Turn 1.22 -> 122%, 0.87 -> 87%, keep strings as-is."""
    try:
        x = float(val)
    except Exception:
        return "" if val is None else str(val)
    # if already looks like 0-3 scale turn to percent
    if -3.0 <= x <= 3.0:
        return f"{x*100:.0f}%"
    return f"{x:.0f}%"

def _replace_in_paragraph(par, mapping: Dict[str, Any]):
    # gather full text across runs
    full = "".join(r.text for r in par.runs)
    # replace on the unified string
    def subfun(m):
        k = m.group(1)
        v = mapping.get(k, "")
        return "" if v is None else str(v)
    new_text = _RX_PH.sub(subfun, full)
    # clear runs and set one new
    for r in list(par.runs):
        r._element.getparent().remove(r._element)
    par.add_run(new_text)

def replace_placeholders_robust(doc: Document, mapping: Dict[str, Any]):
    """Safe replacement in paragraphs + all table cells."""
    for p in doc.paragraphs:
        _replace_in_paragraph(p, mapping)
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    _replace_in_paragraph(p, mapping)

def extract_placeholders_from_docx(doc: Document) -> set[str]:
    """Scan a DOCX and return all [[placeholders]] it contains."""
    found = set()
    def scan(s: str):
        for m in _RX_PH.finditer(s or ""):
            found.add(m.group(1))
    for p in doc.paragraphs:
        scan("".join(r.text for r in p.runs))
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    scan("".join(r.text for r in p.runs))
    return found

def normalize_headers(cols: Iterable[str]) -> list[str]:
    def norm(s: str) -> str:
        s = str(s).strip().lower()
        s = re.sub(r"[^a-z0-9]+", "_", s)  # spaces/greek → underscores
        return s.strip("_")
    return [norm(c) for c in cols]

def col_by_letter(df: pd.DataFrame, letter: str) -> str | None:
    """Map Excel column letter (e.g., 'N', 'AA') to df column name (0-based)."""
    if not letter:
        return None
    L = letter.strip().upper()
    # convert letters to 0-based index
    idx = 0
    for ch in L:
        if not ("A" <= ch <= "Z"):
            return None
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    idx -= 1
    if 0 <= idx < len(df.columns):
        return df.columns[idx]
    return None

def safe_get(row: pd.Series, col: str | None) -> Any:
    if not col or col not in row.index:
        return ""
    v = row[col]
    return "" if pd.isna(v) else v

# ───────────────────────────── UI ─────────────────────────────
st.title("📊 Excel/CSV → 📄 Review/Plan Generator (BEX & Non-BEX)")

left, right = st.columns([1, 1])

with left:
    st.subheader("1) Templates (.docx)")
    tpl_bex = st.file_uploader("BEX template", type=["docx"], key="tpl_bex")
    tpl_non = st.file_uploader("NON-BEX template", type=["docx"], key="tpl_non")
    st.caption("Χρησιμοποίησε placeholders τύπου [[store]], [[plan_vs_target]], [[mobile_actual]] κ.λπ.")

with right:
    st.subheader("2) Excel")
    xls = st.file_uploader("Excel (.xlsx)", type=["xlsx"], key="xls")
    sheet_name = st.text_input("Όνομα φύλλου (Sheet)", value="Sheet1")

st.divider()

with st.expander("Ρυθμίσεις & BEX"):
    debug = st.toggle("🛠 Debug mode", value=False)
    test_mode = st.toggle("🧪 Test mode (πρώτες 50 γραμμές)", value=False)
    st.write("**BEX detection**")
    bex_mode = st.radio("Πως βρίσκουμε αν είναι BEX;", ["Από στήλη (YES/NO)", "Από λίστα κωδικών"], index=0, horizontal=True)
    bex_list_input = st.text_input("BEX λίστα (comma separated)", value="DRZ01,FKM01,ESC01,LND01,PKK01").upper()
    bex_list = set(s.strip() for s in bex_list_input.split(",") if s.strip())

st.subheader("3) Mapping με γράμματα Excel (προαιρετικό)")
map_cols = {}
cols_form = st.columns(4)
labels = [
    ("plan_vs_target", "A"),
    ("mobile_actual", "N"),
    ("mobile_target", "O"),
    ("fixed_target", "P"),
    ("fixed_actual", "Q"),
    ("voice_vs_target", "R"),
    ("fixed_vs_target", "S"),
    ("llu_actual", "T"),
    ("nga_actual", "U"),
    ("ftth_actual", "V"),
    ("eon_tv_actual", "X"),
    ("fwa_actual", "Y"),
    ("mobile_upgrades", "AA"),
    ("fixed_upgrades", "AB"),
    ("pending_mobile", "AF"),
    ("pending_fixed", "AH"),
]
for i, (key, default_letter) in enumerate(labels):
    with cols_form[i % 4]:
        map_cols[key] = st.text_input(key, value=default_letter)

st.divider()
start = st.button("🔧 Generate")

# ───────────────────────────── MAIN ─────────────────────────────
if start:
    # validations
    if xls is None:
        st.error("Ανέβασε Excel πρώτα.")
        st.stop()
    if not tpl_bex or not tpl_non:
        st.error("Ανέβασε και τα δύο templates (.docx).")
        st.stop()

    # read excel
    try:
        xfile = pd.ExcelFile(xls)
        if sheet_name not in xfile.sheet_names:
            st.error(f"Το sheet '{sheet_name}' δεν βρέθηκε. Διαθέσιμα: {xfile.sheet_names}")
            st.stop()
        df_raw = pd.read_excel(xfile, sheet_name=sheet_name)
        df = df_raw.copy()
        df.columns = normalize_headers(df.columns)
    except Exception as e:
        st.error(f"Σφάλμα ανάγνωσης Excel: {e}")
        st.stop()

    # find store column (robust)
    store_col_candidates = ["store_code", "dealer_code", "dealer", "store", "shop_code", "shopcode", "code"]
    store_col = next((c for c in store_col_candidates if c in df.columns), None)
    if not store_col:
        # fallback: first text-like column
        store_col = df.columns[0]

    # attach bex flag
    if bex_mode == "Από στήλη (YES/NO)":
        bex_col_candidates = ["bex", "bex_store", "is_bex", "bex_yes_no"]
        bex_col = next((c for c in bex_col_candidates if c in df.columns), None)
        def _is_bex(row) -> bool:
            val = str(safe_get(row, bex_col)).strip().lower()
            return val in ("yes", "y", "1", "true", "ναι")
    else:
        def _is_bex(row) -> bool:
            return str(safe_get(row, store_col)).strip().upper() in bex_list

    # map Excel letters → normalized df columns
    letter_to_col: Dict[str, str | None] = {k: col_by_letter(df, v) for k, v in map_cols.items()}

    if debug:
        with st.expander("🔎 Mapping preview (letters → headers)"):
            st.json({k: {"letter": map_cols[k], "header": letter_to_col[k]} for k in map_cols})

    # audit templates
    tpl_bex_bytes = tpl_bex.read()
    tpl_non_bytes = tpl_non.read()
    doc_bex = Document(io.BytesIO(tpl_bex_bytes))
    doc_non = Document(io.BytesIO(tpl_non_bytes))
    ph_bex = extract_placeholders_from_docx(doc_bex)
    ph_non = extract_placeholders_from_docx(doc_non)

    with st.expander("🧪 Template audit (placeholders που βρέθηκαν στα .docx)"):
        st.write("BEX template placeholders:", sorted(ph_bex))
        st.write("NON-BEX template placeholders:", sorted(ph_non))

    # generate per row
    out_zip = io.BytesIO()
    zf = zipfile.ZipFile(out_zip, "w", zipfile.ZIP_DEFLATED)

    built = 0
    total_rows = len(df) if not test_mode else min(50, len(df))
    pbar = st.progress(0.0, text="Ξεκίνησε…")

    # Which keys are percentages (format as 122%)
    percent_keys = {"plan_vs_target", "voice_vs_target", "fixed_vs_target"}

    for i, (_, row) in enumerate(df.head(total_rows).iterrows(), start=1):
        try:
            store = str(safe_get(row, store_col)).strip().upper()
            if not store:
                pbar.progress(i/total_rows, text=f"Παράλειψη {i} (κενό store)")
                continue

            is_bex = _is_bex(row)
            tpl_bytes = tpl_bex_bytes if is_bex else tpl_non_bytes

            # build mapping for placeholders
            mapping: Dict[str, Any] = {
                "title": f"Review {TODAY.strftime('%B %Y')} — Plan {(TODAY.replace(day=1) + dt.timedelta(days=32)).strftime('%B %Y')} — {store}",
                "store": store,
                "plan_month": f"Review {TODAY.strftime('%B %Y')} — Plan {(TODAY.replace(day=1) + dt.timedelta(days=32)).strftime('%B %Y')}",
                "bex": "YES" if is_bex else "NO",
            }

            # fill mapped numeric/text fields from letters
            for key, colname in letter_to_col.items():
                val = safe_get(row, colname)
                if key in percent_keys:
                    mapping[key] = format_percent(val)
                else:
                    mapping[key] = "" if val == "" else val

            # also expose every df column as [[<header>]] if user wants it
            for col in df.columns:
                mapping.setdefault(col, safe_get(row, col))

            # create docx
            doc = Document(io.BytesIO(tpl_bytes))
            replace_placeholders_robust(doc, mapping)

            out_name = f"{'BEX' if is_bex else 'NON_BEX'}/{store}_ReviewPlan.docx"
            buf = io.BytesIO()
            doc.save(buf)
            zf.writestr(out_name, buf.getvalue())
            built += 1
            pbar.progress(i/total_rows, text=f"Φτιάχνω: {out_name} ({i}/{total_rows})")
        except Exception as e:
            st.warning(f"⚠️ Σφάλμα στη γραμμή {i}: {e}")

    zf.close()
    pbar.empty()

    if built == 0:
        st.error("Δεν δημιουργήθηκε αρχείο. Έλεγξε templates & mapping.")
    else:
        st.success(f"Έτοιμα {built} αρχεία.")
        st.download_button("⬇️ Κατέβασε ZIP", data=out_zip.getvalue(), file_name="reviews_from_excel.zip")

    if debug:
        with st.expander("🔍 Πρώτη γραμμή (mapping που περάσαμε στο DOCX)"):
            if len(df):
                # δείξε το mapping της πρώτης γραμμής όπως το φτιάχνουμε
                row0 = df.iloc[0]
                sample = {k: (format_percent(safe_get(row0, letter_to_col[k])) if k in percent_keys else safe_get(row0, letter_to_col[k]))
                          for k in letter_to_col}
                sample["store"] = safe_get(row0, store_col)
                st.json(sample)