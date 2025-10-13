# app.py
# Streamlit → Excel/CSV -> (BEX / Non-BEX) Review-Plan .docx (ZIP)
# Placeholders στο .docx: ΔΙΠΛΕΣ αγκύλες (π.χ. [[store]], [[plan_vs_target]])

import io, re, zipfile
from typing import Any, Dict, Optional

import pandas as pd
import streamlit as st
from docx import Document
from docx.oxml.ns import qn

# ───────────── UI CONFIG ─────────────
st.set_page_config(page_title="Excel → Review/Plan Generator", layout="wide")
st.title("📊 Excel/CSV → 📄 Review/Plan Generator (BEX & Non-BEX)")

# ───────────── Helpers ─────────────
PH_RE = re.compile(r"\[\[([A-Za-z0-9_]+)\]\]")

BEX_DEFAULT = {"DRZ01", "FKM01", "ESC01", "LND01", "PKK01"}

def set_default_font(doc: Document, font_name: str = "Aptos") -> None:
    for style in doc.styles:
        if hasattr(style, "font"):
            try:
                style.font.name = font_name
                style._element.rPr.rFonts.set(qn("w:eastAsia"), font_name)
                style._element.rPr.rFonts.set(qn("w:cs"), font_name)
            except Exception:
                pass

def _replace_in_paragraph(par, mapping: Dict[str, Any]):
    full_text = "".join(run.text for run in par.runs) or ""
    if not full_text:
        return
    def repl(m):
        k = m.group(1)
        v = mapping.get(k, "")
        return "" if v is None else str(v)
    new_text = PH_RE.sub(repl, full_text)
    if new_text == full_text:
        return
    # γράψε ενιαίο run ώστε να μην «σπάνε» τα [[...]]
    while len(par.runs) > 1:
        par.runs[-1]._element.getparent().remove(par.runs[-1]._element)
    if par.runs:
        par.runs[0].text = new_text
    else:
        par.add_run(new_text)

def replace_placeholders_everywhere(doc: Document, mapping: Dict[str, Any]):
    # σώμα
    for p in doc.paragraphs:
        _replace_in_paragraph(p, mapping)
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    _replace_in_paragraph(p, mapping)
    # headers/footers
    for section in doc.sections:
        for part in [section.header, section.first_page_header, section.even_page_header,
                     section.footer, section.first_page_footer, section.even_page_footer]:
            if not part:
                continue
            for p in part.paragraphs:
                _replace_in_paragraph(p, mapping)
            for t in part.tables:
                for row in t.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            _replace_in_paragraph(p, mapping)

def normkey(x: Any) -> str:
    return re.sub(r"[\s\-_\.]+", "", str(x).strip().lower())

def pick(columns, *aliases) -> Optional[str]:
    nmap = {normkey(c): c for c in columns}
    for a in aliases:
        if normkey(a) in nmap:
            return nmap[normkey(a)]
    for a in aliases:
        pat = re.compile(a, re.IGNORECASE)
        for c in columns:
            if re.search(pat, str(c)):
                return c
    return None

def read_data(xls, sheet_name: str) -> Optional[pd.DataFrame]:
    try:
        name = getattr(xls, "name", "")
        if name.lower().endswith(".csv"):
            st.write("📑 Sheets:", ["CSV Data"])
            return pd.read_csv(xls)
        xfile = pd.ExcelFile(xls, engine="openpyxl")
        st.write("📑 Sheets:", xfile.sheet_names)
        if sheet_name not in xfile.sheet_names:
            st.error(f"Το sheet '{sheet_name}' δεν βρέθηκε. Διαθέσιμα: {xfile.sheet_names}")
            return None
        return pd.read_excel(xfile, sheet_name=sheet_name, engine="openpyxl")
    except Exception as e:
        st.error(f"Δεν άνοιξε το αρχείο: {e}")
        return None

# Excel letter helpers
def letter_to_index(s: str) -> Optional[int]:
    """A -> 0, B -> 1, ..., Z -> 25, AA -> 26, ...  Επιστρέφει None αν είναι κενό."""
    if not s:
        return None
    s = s.strip().upper()
    if not re.fullmatch(r"[A-Z]+", s):
        return None
    n = 0
    for ch in s:
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n - 1

def val_by_letter(row: pd.Series, letter: str):
    idx = letter_to_index(letter)
    if idx is None:
        return ""
    try:
        v = row.iloc[idx]
        return "" if pd.isna(v) else v
    except Exception:
        return ""

def fmt_percent(x, decimals: int = 0):
    """1.22 -> 122% (0 δεκαδικά default). Αν είναι ήδη %, το σέβεται."""
    if x is None or x == "":
        return ""
    try:
        # Αν έρθει string με % ή , ως δεκαδικό
        s = str(x).strip()
        if s.endswith("%"):
            return s
        s = s.replace(",", ".")
        val = float(s)
        return f"{round(val*100, decimals):.{decimals}f}%" if decimals > 0 else f"{round(val*100):d}%"
    except Exception:
        return str(x)

# ───────────── Sidebar ─────────────
with st.sidebar:
    st.header("Ρυθμίσεις")
    debug_mode = st.toggle("🛠 Debug mode", value=True)
    test_mode  = st.toggle("🧪 Test mode (πρώτες 50 γραμμές)", value=True)

    st.markdown("### Templates (.docx)")
    tpl_bex    = st.file_uploader("BEX template", type=["docx"])
    tpl_nonbex = st.file_uploader("Non-BEX template", type=["docx"])
    st.caption("Placeholders: [[title]], [[plan_month]], [[store]], [[bex]], "
               "[[plan_vs_target]], [[mobile_actual]], [[mobile_target]], [[fixed_target]], "
               "[[fixed_actual]], [[voice_vs_target]], [[fixed_vs_target]], [[llu_actual]], "
               "[[nga_actual]], [[ftth_actual]], [[eon_tv_actual]], [[fwa_actual]], "
               "[[mobile_upgrades]], [[fixed_upgrades]], [[pending_mobile]], [[pending_fixed]]")

    st.markdown("### STORE & BEX")
    bex_mode = st.radio("Πώς βρίσκουμε αν είναι BEX;", ["Σταθερή λίστα (DRZ01, FKM01, ESC01, LND01, PKK01)", "Από στήλη (YES/NO)"], index=0)
    bex_list_text = st.text_input("Σταθερή λίστα (comma-sep)", "DRZ01, FKM01, ESC01, LND01, PKK01")
    bex_list = {s.strip().upper() for s in bex_list_text.split(",") if s.strip()}
    bex_letter = st.text_input("Γράμμα στήλης για BEX (YES/NO) [προαιρετικό]", "")

    st.markdown("### Mapping με γράμματα Excel (A, N, AA, AB, AF, AH)")
    letter_store       = st.text_input("Store (αν ΔΕΝ βρεθεί από header)", "")
    letter_plan_vs     = st.text_input("plan_vs_target", "A")
    letter_mobile_act  = st.text_input("mobile_actual", "N")
    letter_mobile_tgt  = st.text_input("mobile_target", "O")
    letter_fixed_tgt   = st.text_input("fixed_target", "P")
    letter_fixed_act   = st.text_input("fixed_actual", "Q")
    letter_voice_vs    = st.text_input("voice_vs_target", "R")
    letter_fixed_vs    = st.text_input("fixed_vs_target", "S")
    letter_llu         = st.text_input("llu_actual", "T")
    letter_nga         = st.text_input("nga_actual", "U")
    letter_ftth        = st.text_input("ftth_actual", "V")
    letter_eon_tv      = st.text_input("eon_tv_actual", "X")
    letter_fwa         = st.text_input("fwa_actual", "Y")
    letter_mob_upg     = st.text_input("mobile_upgrades", "AA")
    letter_fix_upg     = st.text_input("fixed_upgrades", "AB")
    letter_pend_mob    = st.text_input("pending_mobile", "AF")
    letter_pend_fix    = st.text_input("pending_fixed", "AH")

# ───────────── Main inputs ─────────────
st.markdown("### 1) Ανέβασε Excel/CSV")
xls = st.file_uploader("Excel ή CSV", type=["xlsx", "csv"])
sheet_name = st.text_input("Όνομα φύλλου (Excel)", value="Sheet1")
run = st.button("🔧 Generate")

# ───────────── MAIN ─────────────
if run:
    if not xls:
        st.error("Ανέβασε αρχείο Excel/CSV.")
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
        st.dataframe(df.head(8))

    # Auto pick headers για store αν υπάρχουν
    col_store = pick(df.columns, "Dealer_Code", "Dealer Code", "Shop Code", "Shop_Code", "ShopCode", "Κατάστημα", r"shop.?code")
    if debug_mode:
        st.write("STORE από header:", col_store or "(none)")

    # Προετοιμασία templates
    tpl_bex_bytes    = tpl_bex.read()
    tpl_nonbex_bytes = tpl_nonbex.read()

    out_zip = io.BytesIO()
    zf = zipfile.ZipFile(out_zip, "w", zipfile.ZIP_DEFLATED)
    built = 0

    # Έλεγχος πλήθους
    total = len(df) if not test_mode else min(50, len(df))
    pbar = st.progress(0, text="Δημιουργία εγγράφων…")

    for i, (_, row) in enumerate(df.iterrows(), start=1):
        if test_mode and i > total:
            st.info(f"🧪 Test mode: σταμάτησα στις {total} γραμμές.")
            break

        # --- Τιμές από header ή γράμμα ---
        store_val = (str(row[col_store]).strip() if col_store else "") or str(val_by_letter(row, letter_store)).strip()
        if not store_val:
            pbar.progress(min(i/(total or 1), 1.0), text=f"Παράλειψη γραμμής {i} (κενό store)")
            continue
        store_up = store_val.upper()

        # BEX flag
        if bex_mode.startswith("Σταθερή"):
            is_bex = store_up in (bex_list or BEX_DEFAULT)
        else:
            raw_bex = str(val_by_letter(row, bex_letter)).strip().lower()
            is_bex = raw_bex in {"yes", "y", "1", "true", "ναι"}

        # Λήψη πεδίων
        plan_vs_target  = val_by_letter(row, letter_plan_vs)
        mobile_actual   = val_by_letter(row, letter_mobile_act)
        mobile_target   = val_by_letter(row, letter_mobile_tgt)
        fixed_target    = val_by_letter(row, letter_fixed_tgt)
        fixed_actual    = val_by_letter(row, letter_fixed_act)
        voice_vs_target = val_by_letter(row, letter_voice_vs)
        fixed_vs_target = val_by_letter(row, letter_fixed_vs)
        llu_actual      = val_by_letter(row, letter_llu)
        nga_actual      = val_by_letter(row, letter_nga)
        ftth_actual     = val_by_letter(row, letter_ftth)
        eon_tv_actual   = val_by_letter(row, letter_eon_tv)
        fwa_actual      = val_by_letter(row, letter_fwa)
        mobile_upgrades = val_by_letter(row, letter_mob_upg)
        fixed_upgrades  = val_by_letter(row, letter_fix_upg)
        pending_mobile  = val_by_letter(row, letter_pend_mob)
        pending_fixed   = val_by_letter(row, letter_pend_fix)

        # Μορφοποίηση ποσοστών (1.22 -> 122%)
        plan_vs_target_fmt  = fmt_percent(plan_vs_target)
        voice_vs_target_fmt = fmt_percent(voice_vs_target)
        fixed_vs_target_fmt = fmt_percent(fixed_vs_target)

        mapping = {
            "title":      f"Review September 2025 — Plan October 2025 — {store_up}",
            "plan_month": "Review September 2025 → Plan October 2025",
            "store":      store_up,
            "bex":        "YES" if is_bex else "NO",

            "plan_vs_target":  plan_vs_target_fmt,
            "mobile_actual":   mobile_actual,
            "mobile_target":   mobile_target,
            "fixed_target":    fixed_target,
            "fixed_actual":    fixed_actual,
            "voice_vs_target": voice_vs_target_fmt,
            "fixed_vs_target": fixed_vs_target_fmt,

            "llu_actual":      llu_actual,
            "nga_actual":      nga_actual,
            "ftth_actual":     ftth_actual,
            "eon_tv_actual":   eon_tv_actual,
            "fwa_actual":      fwa_actual,

            "mobile_upgrades": mobile_upgrades,
            "fixed_upgrades":  fixed_upgrades,
            "pending_mobile":  pending_mobile,
            "pending_fixed":   pending_fixed,
        }

        try:
            doc = Document(io.BytesIO(tpl_bex_bytes if is_bex else tpl_nonbex_bytes))
            set_default_font(doc, "Aptos")
            replace_placeholders_everywhere(doc, mapping)

            out_name = f"{store_up}_ReviewSep_PlanOct.docx"
            buf = io.BytesIO()
            doc.save(buf)
            zf.writestr(out_name, buf.getvalue())
            built += 1
            pbar.progress(min(i/(total or 1), 1.0), text=f"Φτιάχνω: {out_name} ({min(i,total)}/{total})")
        except Exception as e:
            st.warning(f"⚠️ Γραμμή {i}: {e}")
            if debug_mode:
                st.exception(e)

    zf.close()
    pbar.empty()

    if built == 0:
        st.error("Δεν δημιουργήθηκε αρχείο. Έλεγξε STORE mapping & templates.")
    else:
        st.success(f"Έτοιμα {built} αρχεία.")
        st.download_button("⬇️ Κατέβασε ZIP", data=out_zip.getvalue(), file_name="reviews_from_excel.zip")