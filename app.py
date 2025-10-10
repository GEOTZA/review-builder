import streamlit as st
import io, zipfile, re
import pandas as pd
from typing import Dict, Any
from docx import Document
from docx.oxml.ns import qn

st.set_page_config(page_title="Excel → Review/Plan Generator", layout="wide")

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

def normkey(x: str) -> str:
    """Κανονικοποίηση header: πεζά, χωρίς κενά/underscores/τελείες/παύλες."""
    return re.sub(r"[\s\-_\.]+", "", str(x).strip().lower())

def pick(columns, *aliases):
    """Βρες στήλη με βάση aliases (normalized)."""
    nmap = {normkey(c): c for c in columns}
    for a in aliases:
        if normkey(a) in nmap:
            return nmap[normkey(a)]
    # βρες με περιεχόμενο (contains)
    for a in aliases:
        pat = re.compile(a, re.IGNORECASE)
        for c in columns:
            if re.search(pat, str(c)):
                return c
    return None

# ---------- UI ----------
st.title("📊 Excel → 📄 Review/Plan Generator (BEX & Non-BEX)")

with st.sidebar:
    st.header("⚙️ BEX")
    bex_mode = st.radio("Πηγή BEX", ["Στήλη στο Excel", "Λίστα (comma-separated)"], index=0)
    bex_list = set()
    if bex_mode == "Λίστα (comma-separated)":
        bex_input = st.text_area("BEX stores", "ESC01,FKM01,LND01,DRZ01,PKK01")
        bex_list = set(s.strip().upper() for s in bex_input.split(",") if s.strip())

    st.subheader("📄 Templates (.docx)")
    tpl_bex = st.file_uploader("BEX template", type=["docx"])
    tpl_nonbex = st.file_uploader("Non-BEX template", type=["docx"])
    st.caption("Placeholders: [[title]], [[store]], [[mobile_actual]], [[mobile_target]], [[fixed_actual]], [[fixed_target]], [[pending_mobile]], [[pending_fixed]], [[plan_vs_target]]")

st.markdown("### 1) Ανέβασε Excel")
xls = st.file_uploader("Excel (xlsx)", type=["xlsx"])
sheet_name = st.text_input("Όνομα φύλλου (Sheet)", value="Sheet1")

run = st.button("🔧 Generate")

if run:
    if not xls:
        st.error("Ανέβασε Excel πρώτα.")
        st.stop()
    if not tpl_bex or not tpl_nonbex:
        st.error("Ανέβασε και τα δύο templates.")
        st.stop()

    try:
        df = pd.read_excel(xls, sheet_name=sheet_name)
    except Exception as e:
        st.error(f"Δεν άνοιξε το Excel (sheet '{sheet_name}'): {e}")
        st.stop()
    if df.empty:
        st.error("Το Excel είναι άδειο.")
        st.stop()

    cols = list(df.columns)

    # ---- AUTO-MAP με βάση το screenshot σου ----
    col_store       = pick(cols, "Shop Code", "Shop_Code", "ShopCode", "STORE", "Κατάστημα", r"shop.?code")
    col_bex         = pick(cols, "BEX store", "BEX", r"bex.?store")
    col_mob_act     = pick(cols, "mobile actual", r"mobile.*actual")
    col_mob_tgt     = pick(cols, "mobile target", r"mobile.*target", "mobile plan")
    col_fix_tgt     = pick(cols, "target fixed", r"fixed.*target", "fixed plan total", "fixed plan")
    col_fix_act     = pick(cols, "total fixed", r"(total|sum).?fixed.*actual", "fixed actual")
    col_pend_mob    = pick(cols, "TOTAL PENDING MOBILE", r"pending.*mobile")
    col_pend_fix    = pick(cols, "TOTAL PENDING FIXED", r"pending.*fixed")
    col_plan_vs     = pick(cols, "plan vs target", r"plan.*vs.*target")

    # προβολή που βρήκαμε
    with st.expander("Χαρτογράφηση (auto)"):
        st.write({
            "STORE": col_store, "BEX": col_bex,
            "mobile_actual": col_mob_act, "mobile_target": col_mob_tgt,
            "fixed_target": col_fix_tgt, "fixed_actual": col_fix_act,
            "pending_mobile": col_pend_mob, "pending_fixed": col_pend_fix,
            "plan_vs_target": col_plan_vs
        })

    tpl_bex_bytes = tpl_bex.read()
    tpl_nonbex_bytes = tpl_nonbex.read()

    out_zip = io.BytesIO()
    z = zipfile.ZipFile(out_zip, "w", zipfile.ZIP_DEFLATED)
    built = 0

    def cell(row, col):
        if not col: return ""
        v = row[col]
        if pd.isna(v): return ""
        return v

    for _, row in df.iterrows():
        store = str(cell(row, col_store)).strip()
        if not store:
            continue
        store_up = store.upper()

        # BEX flag
        if bex_mode == "Λίστα (comma-separated)":
            is_bex = store_up in bex_list
        else:
            bex_val = str(cell(row, col_bex)).strip().lower()
            is_bex = bex_val in ("yes","y","1","true","ναι")

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

        # Build doc
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
        st.error("Δεν δημιουργήθηκε αρχείο. Έλεγξε αν αναγνωρίστηκε η στήλη STORE (π.χ. 'Shop Code').")
    else:
        st.success(f"Έτοιμα {built} αρχεία.")
        st.download_button("⬇️ Κατέβασε ZIP", data=out_zip.getvalue(), file_name="reviews_from_excel.zip")
