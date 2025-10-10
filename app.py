import streamlit as st
import io, zipfile, re, json
import pandas as pd
from typing import Dict, Any
from docx import Document
from docx.oxml.ns import qn

# PDF text extraction
try:
    from pdfminer.high_level import extract_text
except Exception:
    extract_text = None

st.set_page_config(page_title="Excel KPIs + Review (PDF/Manual)", layout="wide")

# ---------- helpers ----------
def set_default_font(doc, font_name="Aptos"):
    for style in doc.styles:
        if hasattr(style, "font"):
            try:
                style.font.name = font_name
                style._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
                style._element.rPr.rFonts.set(qn('w:cs'), font_name)
            except Exception:
                pass

def replace_placeholders(doc: Document, mapping: Dict[str, Any]):
    def repl_text(s: str) -> str:
        def rfun(m):
            k = m.group(1)
            v = mapping.get(k, "")
            return "" if v is None else str(v)
        return re.sub(r"\[\[([A-Za-z0-9_]+)\]\]", rfun, s)
    for p in doc.paragraphs:
        for r in p.runs:
            r.text = repl_text(r.text)
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        r.text = repl_text(r.text)

def read_pdf_text(name: str, data: bytes) -> str:
    if extract_text is None: 
        return ""
    try:
        return extract_text(io.BytesIO(data)) or ""
    except Exception:
        return ""

def guess_col(columns, *keys):
    """Ευρετική: βρίσκει 1η στήλη που ταιριάζει σε keywords."""
    for k in keys:
        pat = re.compile(k, re.IGNORECASE)
        for c in columns:
            if re.search(pat, str(c)):
                return c
    return None

def parse_file_role(name: str):
    """
    Χαλαρό matching για filenames:
      <STORE>_Business Review.pdf, <STORE>_Action Plan.pdf
      ή ESC01_Review.pdf / ESC01_Plan.pdf κ.λπ.
    """
    base = name.strip()
    if base.lower().endswith(".pdf"):
        base = base[:-4]
    base = re.sub(r"\s+", " ", base)
    m_store = re.match(r"^\s*([A-Za-z0-9]+)[\s_\-]+(.+)$", base)
    if not m_store:
        only_code = re.match(r"^\s*([A-Za-z0-9]+)\s*$", base)
        if only_code:
            return only_code.group(1).upper(), None
        return None, None
    store = m_store.group(1).upper()
    tail = m_store.group(2).strip().lower()
    compact = re.sub(r"[\s_\-]+", "", tail)
    if any(k in compact for k in ["businessreview","review"]):
        return store, "review"
    if any(k in compact for k in ["actionplan","plan"]):
        return store, "plan"
    if "review" in compact: return store, "review"
    if "plan" in compact:   return store, "plan"
    return store, None

# ---------- UI ----------
st.title("📊 KPIs από Excel + 📝 Review από PDF/Manual")

with st.sidebar:
    st.header("⚙️ Ρυθμίσεις")
    bex_mode = st.radio("BEX ορισμός", ["Λίστα (comma-separated)", "Στήλη στο Excel"], index=0)
    bex_list = set()
    bex_col = None
    if bex_mode == "Λίστα (comma-separated)":
        bex_input = st.text_area("BEX stores", "ESC01,FKM01,LND01,DRZ01,PKK01")
        bex_list = set([s.strip().upper() for s in bex_input.split(",") if s.strip()])

    st.subheader("📄 Templates (.docx)")
    tpl_bex = st.file_uploader("BEX template", type=["docx"], key="tpl_bex")
    tpl_nonbex = st.file_uploader("Non-BEX template", type=["docx"], key="tpl_nonbex")
    st.caption("Placeholders: [[title]], [[store]], [[plan_month]], [[mobile_actual]], [[fixed_actual]], [[mobile_target]], [[fixed_target]], [[review_body]]")

st.markdown("### 1) Ανέβασε Excel (KPIs)")
xls = st.file_uploader("Excel (xlsx)", type=["xlsx"])
sheet_name = st.text_input("Sheet name", value="Sheet1")

st.markdown("### 2) Ανέβασε τα Business Review PDFs (προαιρετικά)")
st.caption("Φόρτωσε αρχεία τύπου <STORE>_Business Review.pdf (ή ESC01_Review.pdf κ.λπ.). Αν δεν έχει κείμενο/λείπει, θα μπορείς να γράψεις manual review.")
uploaded_pdfs = st.file_uploader("PDFs", type=["pdf"], accept_multiple_files=True)

st.markdown("### 3) Λίστα καταστημάτων")
store_list_text = st.text_area("STORE codes (ένα ανά γραμμή)", "ESC01\nLCI01\nLGS01\nLND01")

run = st.button("🔧 Generate DOCX")

if run:
    # --- validations
    if not xls:
        st.error("Ανέβασε Excel πρώτα.")
        st.stop()
    if not tpl_bex or not tpl_nonbex:
        st.error("Ανέβασε και τα δύο templates (.docx).")
        st.stop()

    # --- read Excel
    try:
        df = pd.read_excel(xls, sheet_name=sheet_name)
    except Exception as e:
        st.error(f"Αποτυχία ανάγνωσης Excel: {e}")
        st.stop()
    if df.empty:
        st.error("Το Excel είναι άδειο.")
        st.stop()

    cols = list(df.columns)
    # auto-guess για τα 6 βασικά
    col_store       = guess_col(cols, r"shop|store|dealer|code|shop_code|store_code|κατάστημα")
    col_mobile_act  = guess_col(cols, r"^mobile.*actual$|mobileactual|κινητ.*actual|κινητ.*παραγ")
    col_fixed_act   = guess_col(cols, r"^total.*fixed.*actual$|fixed.*actual|σταθερ.*actual|σταθερ.*παραγ")
    col_mobile_tgt  = guess_col(cols, r"^mobile.*(target|plan)$|κινητ.*(στόχος|πλάνο)")
    col_fixed_tgt   = guess_col(cols, r"^fixed.*(target|plan)$|σταθερ.*(στόχος|πλάνο)")
    col_plan_month  = guess_col(cols, r"^plan.*month$|μήνας.*πλάνου|μηνας.*πλανου")
    if bex_mode == "Στήλη στο Excel":
        bex_col = guess_col(cols, r"^bex$|bex store|is_bex|bex_yes_no")

    with st.expander("Χαρτογράφηση πεδίων (auto-guess)"):
        col_store      = st.selectbox("STORE", options=cols, index=(cols.index(col_store) if col_store in cols else 0))
        if bex_mode == "Στήλη στο Excel":
            bex_col = st.selectbox("BEX column", options=["(none)"] + cols, index=((cols.index(bex_col)+1) if bex_col in cols else 0))
            if bex_col == "(none)": bex_col = None
        col_mobile_act = st.selectbox("Mobile Actual", options=["(none)"] + cols, index=((cols.index(col_mobile_act)+1) if col_mobile_act in cols else 0))
        col_fixed_act  = st.selectbox("Fixed Actual",  options=["(none)"] + cols, index=((cols.index(col_fixed_act)+1) if col_fixed_act  in cols else 0))
        col_mobile_tgt = st.selectbox("Mobile Target", options=["(none)"] + cols, index=((cols.index(col_mobile_tgt)+1) if col_mobile_tgt in cols else 0))
        col_fixed_tgt  = st.selectbox("Fixed Target",  options=["(none)"] + cols, index=((cols.index(col_fixed_tgt)+1) if col_fixed_tgt  in cols else 0))
        col_plan_month = st.selectbox("Plan Month (optional)", options=["(none)"] + cols, index=((cols.index(col_plan_month)+1) if col_plan_month in cols else 0))

    # --- index PDFs
    pdf_map = {}
    if uploaded_pdfs:
        for f in uploaded_pdfs:
            store, role = parse_file_role(f.name)
            if store and (role == "review"):
                pdf_map[store] = f.read()

    # --- load templates
    tpl_bex_bytes = tpl_bex.read()
    tpl_nonbex_bytes = tpl_nonbex.read()

    # --- build all
    out_zip = io.BytesIO()
    z = zipfile.ZipFile(out_zip, "w", zipfile.ZIP_DEFLATED)

    stores = [s.strip().upper() for s in store_list_text.splitlines() if s.strip()]
    built = 0

    for _, row in df.iterrows():
        store = str(row[col_store]).strip() if not pd.isna(row[col_store]) else ""
        if not store or store.upper() not in stores:
            continue
        store_up = store.upper()

        # BEX flag
        is_bex = False
        if bex_mode == "Λίστα (comma-separated)":
            is_bex = store_up in bex_list
        else:
            if bex_col and (bex_col in df.columns):
                val = str(row[bex_col]).strip().lower() if not pd.isna(row[bex_col]) else ""
                is_bex = val in ("1","yes","true","y","bex","ναι")

        # KPIs from Excel
        def val(col):
            if not col or col == "(none)": return ""
            v = row[col]
            return "" if pd.isna(v) else v

        mobile_actual = val(col_mobile_act)
        fixed_actual  = val(col_fixed_act)
        mobile_target = val(col_mobile_tgt)
        fixed_target  = val(col_fixed_tgt)
        plan_month    = val(col_plan_month)

        # REVIEW BODY: PDF → κείμενο ή Manual fallback
        review_text = ""
        if store_up in pdf_map:
            txt = read_pdf_text(pdf_map[store_up] and f"{store_up}_BR.pdf", pdf_map[store_up])
            if txt and len(txt.strip()) >= 20:
                # basic καθάρισμα
                review_text = re.sub(r"\n{3,}", "\n\n", txt).strip()
        if not review_text:
            # αν δεν βρέθηκε/ήταν image, δείξε textarea για manual
            with st.expander(f"✍️ Review body για {store_up} (PDF δεν έδωσε κείμενο)"):
                review_text = st.text_area(f"{store_up} review body", "", key=f"rv_{store_up}")

        mapping = {
            "title": f"Review September 2025 — Plan October 2025 — {store_up}",
            "store": store_up,
            "plan_month": plan_month,
            "mobile_actual": mobile_actual,
            "fixed_actual": fixed_actual,
            "mobile_target": mobile_target,
            "fixed_target": fixed_target,
            "review_body": review_text or "",
        }

        doc = Document(io.BytesIO(tpl_bex_bytes if is_bex else tpl_nonbex_bytes))
        set_default_font(doc, "Aptos")
        replace_placeholders(doc, mapping)

        out_name = f"{store_up}_ReviewSep_PlanOct.docx"
        buf = io.BytesIO()
        doc.save(buf)
        z.writestr(out_name, buf.getvalue())
        built += 1

    z.close()
    if built == 0:
        st.error("Δεν δημιουργήθηκε αρχείο. Έλεγξε STORE mapping & λίστα καταστημάτων.")
    else:
        st.success(f"Ολοκληρώθηκε: {built} αρχεία.")
        st.download_button("⬇️ ZIP", data=out_zip.getvalue(), file_name="reviews_from_excel.zip")
