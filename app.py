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
    """Πεζά + αφαίρεση κενών/underscores/παυλών/τελειών, για robust matching."""
    return re.sub(r"[\s\-_\.]+", "", str(x).strip().lower())

def pick(columns, *aliases):
    """Βρες στήλη με βάση aliases (normalized). Πρώτα exact normalized, μετά contains regex."""
    nmap = {normkey(c): c for c in columns}
    # exact normalized
    for a in aliases:
        if normkey(a) in nmap:
            return nmap[normkey(a)]
    # contains pattern
    for a in aliases:
        pat = re.compile(a, re.IGNORECASE)
        for c in columns:
            if re.search(pat, str(c)):
                return c
    return None

def cell(row, col):
    if not col:
        return ""
    # Το 'row' μπορεί να είναι Series ή tuple, το row[col] προσπελαύνει την τιμή
    v = row[col]
    return "" if pd.isna(v) else v

# ---------- UI ----------
st.title("📊 Excel/CSV → 📄 Review/Plan Generator (BEX & Non-BEX)")
debug_mode = st.sidebar.toggle("🛠 Debug mode", value=True)
test_mode  = st.sidebar.toggle("🧪 Test mode (limit rows=50)", value=True)

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

st.markdown("### 1) Ανέβασε Excel/CSV")
xls = st.file_uploader("Excel/CSV", type=["xlsx", "csv"])
sheet_name = st.text_input("Όνομα φύλλου (Sheet - μόνο για Excel)", value="Sheet1")

run = st.button("🔧 Generate")
def load_df_from_excel(xls, sheet_name: str) -> pd.DataFrame:
    xfile = pd.ExcelFile(xls, engine="openpyxl")
    if sheet_name not in xfile.sheet_names:
        raise ValueError(f"Το sheet '{sheet_name}' δεν βρέθηκε. Διαθέσιμα: {xfile.sheet_names}")
    return pd.read_excel(xfile, sheet_name=sheet_name, engine="openpyxl")
if run:
    # 1. Βήμα: Αρχικοί έλεγχοι αρχείων
    if not xls:
        st.error("Ανέβασε αρχείο Excel ή CSV πρώτα.")
        st.stop()
    if not tpl_bex or not tpl_nonbex:
        st.error("Ανέβασε και τα δύο templates.")
        st.stop()

    st.info(
        f"📄 Δεδομένα: {len(xls.getbuffer())/1024:.1f} KB | "
        f"BEX tpl: {tpl_bex.size/1024:.1f} KB | Non-BEX tpl: {tpl_nonbex.size/1024:.1f} KB"
    )
    
    file_type = xls.name.split('.')[-1].lower()
    df = None # Αρχικοποίηση

    # 2. Βήμα: Ανάγνωση αρχείου και έλεγχος sheets
   # 2) Δείξε διαθέσιμα sheets & διάβασε με openpyxl

    try:
        df = load_df_from_excel(xls, sheet_name)
    except Exception as e:
        st.error(f"Δεν άνοιξε το Excel: {e}")
        st.stop()

                df = pd.read_excel(xfile, sheet_name=sheet_name, engine="openpyxl")
            else:
                st.error("Μη υποστηριζόμενος τύπος αρχείου.")
                st.stop()

        except Exception as e:
            st.error(f"Δεν άνοιξε το αρχείο: {e}")
            st.stop()
            
    # --- Ο ΚΩΔΙΚΑΣ ΕΔΩ ΕΚΤΕΛΕΙΤΑΙ ΜΟΝΟ ΑΝ ΤΟ df ΔΙΑΒΑΣΤΗΚΕ ΕΠΙΤΥΧΩΣ ---
    
    if df is None:
        st.error("Αδυναμία φόρτωσης δεδομένων.")
        st.stop()


    st.success(f"OK: {len(df)} γραμμές, {len(df.columns)} στήλες.")
    if debug_mode:
        st.dataframe(df.head(10))

    cols = list(df.columns)

    # ---- AUTO-MAP βασισμένο στο Excel σου ----
    col_store       = pick(cols, "Shop Code", "Shop_Code", "ShopCode", "Shop code", "STORE", "Κατάστημα", r"shop.?code")
    col_bex         = pick(cols, "BEX store", "BEX", r"bex.?store")
    col_mob_act     = pick(cols, "mobile actual", r"mobile.*actual")
    col_mob_tgt     = pick(cols, "mobile target", r"mobile.*target", "mobile plan")
    col_fix_tgt     = pick(cols, "target fixed", r"fixed.*target", "fixed plan total", "fixed plan")
    col_fix_act     = pick(cols, "total fixed", r"(total|sum).?fixed.*actual", "fixed actual")
    col_pend_mob    = pick(cols, "TOTAL PENDING MOBILE", r"pending.*mobile")
    col_pend_fix    = pick(cols, "TOTAL PENDING FIXED", r"pending.*fixed")
    col_plan_vs     = pick(cols, "plan vs target", r"plan.*vs.*target")

    with st.expander("Χαρτογράφηση (auto)"):
        st.write({
            "STORE": col_store, 
            "BEX": col_bex,
            "mobile_actual": col_mob_act, 
            "mobile_target": col_mob_tgt,
            "fixed_target": col_fix_tgt,
            "fixed_actual": col_fix_act,
            "pending_mobile": col_pend_mob, 
            "pending_fixed": col_pend_fix,
            "plan_vs_target": col_plan_vs
        })

    if not col_store:
        st.error("Δεν βρέθηκε στήλη STORE (π.χ. 'Shop Code'). Διόρθωσε την κεφαλίδα ή πρόσθεσε alias.")
        st.stop()

    tpl_bex_bytes = tpl_bex.read()
    tpl_nonbex_bytes = tpl_nonbex.read()

    out_zip = io.BytesIO()
    z = zipfile.ZipFile(out_zip, "w", zipfile.ZIP_DEFLATED)
    built = 0 # Η μεταβλητή για τον μετρητή των αρχείων.

    pbar = st.progress(0, text="Δημιουργία εγγράφων...")
    
    max_rows = 50 if test_mode else len(df)
    
    for i, row_tuple in enumerate(df.itertuples(index=False), start=1):
        if i > max_rows:
            if debug_mode:
                st.info(f"🧪 Test mode: σταμάτησα στις {max_rows} γραμμές.")
            break
        try:
            # Μετατροπή του tuple σε Series για να δουλέψει σωστά το cell(...)
            row = pd.Series(row_tuple, index=df.columns) 
            store = str(cell(row, col_store)).strip()
            
            if not store:
                # Χρησιμοποιούμε max_rows για τον υπολογισμό της προόδου
                pbar.progress(min(i/max_rows, 1.0), text=f"Παράλειψη γραμμής {i} (κενό store)")
                continue

            store_up = store.upper()
            if bex_mode == "Λίστα (comma-separated)":
                is_bex = store_up in bex_list
            else:
                # Χρησιμοποιούμε col_bex μόνο αν το αρχείο είναι Excel 
                # Για CSV, αν δεν υπάρχει col_bex, θεωρούμε ότι δεν είναι BEX (ή το αντίθετο αν το default άλλαζε)
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

            tpl_bytes = tpl_bex_bytes if is_bex else tpl_nonbex_bytes
            doc = Document(io.BytesIO(tpl_bytes))
            set_default_font(doc, "Aptos")
            replace_placeholders(doc, mapping)

            out_name = f"{store_up}_ReviewSep_PlanOct.docx"
            buf = io.BytesIO()
            doc.save(buf)
            z.writestr(out_name, buf.getvalue())
            
            # --- Αυξάνουμε τον μετρητή επιτυχημένων αρχείων ---
            built += 1 

            pbar.progress(min(i/max_rows, 1.0), text=f"Φτιάχνω: {out_name} ({i}/{max_rows})")
        except Exception as e:
            st.warning(f"⚠️ Σφάλμα στη γραμμή {i}: {e}")
            if debug_mode:
                st.exception(e)

    # --- ΤΟ ΤΕΛΟΣ ΤΟΥ if run: BLOCK ---
    z.close()
    if built == 0:
        st.error("Δεν δημιουργήθηκε αρχείο. Έλεγξε STORE mapping & templates.")
    else:
        st.success(f"Έτοιμα {built} αρχεία.")
        st.download_button("⬇️ Κατέβασε ZIP", data=out_zip.getvalue(), file_name="reviews_from_excel.zip")
