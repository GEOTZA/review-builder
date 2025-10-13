# app.py
# Streamlit: Excel/CSV → (BEX / Non-BEX) Review-Plan .docx (ZIP)

import io
import re
import zipfile
from typing import Any, Dict, Optional

import pandas as pd
import streamlit as st
from docx import Document
from docx.oxml.ns import qn
from openpyxl.utils import column_index_from_string

# ───────────────────────────── UI CONFIG ─────────────────────────────
st.set_page_config(page_title="Excel → Review/Plan Generator", layout="wide")
st.title("📚 Excel/CSV → 📄 Review/Plan Generator (BEX & Non-BEX)")

# ───────────────────────────── HELPERS ─────────────────────────────
def set_default_font(doc: Document, font_name: str = "Aptos") -> None:
    """Default font σε όλα τα styles (και eastAsia)."""
    for style in doc.styles:
        if hasattr(style, "font"):
            try:
                style.font.name = font_name
                style._element.rPr.rFonts.set(qn("w:eastAsia"), font_name)
                style._element.rPr.rFonts.set(qn("w:cs"), font_name)
            except Exception:
                pass

def replace_placeholders(doc: Document, mapping: Dict[str, Any]) -> None:
    """Αντικαθιστά [[placeholders]] σε paragraphs & tables."""
    pattern = re.compile(r"\[\[([A-Za-z0-9_]+)\]\]")

    def subfun(s: str) -> str:
        key_to_val = lambda m: "" if mapping.get(m.group(1)) is None else str(mapping.get(m.group(1), ""))
        return pattern.sub(key_to_val, s)

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
    """lower + αφαίρεση κενών/-,_,. για robust ταύτιση headers."""
    return re.sub(r"[\s\-_\.]+", "", str(x).strip().lower())

def pick(columns, *aliases) -> str:
    """Βρες στήλη με βάση aliases (πρώτα exact normalized, μετά regex contains)."""
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

def col_from_letter(letter: Optional[str]) -> Optional[int]:
    """'AA' -> 27 (1-based)."""
    if not letter:
        return None
    try:
        return int(column_index_from_string(letter.strip().upper()))
    except Exception:
        return None

def get_cell_by_letter(df: pd.DataFrame, letter: Optional[str], row_index_zero_based: int) -> tuple[str, str]:
    """
    Επιστρέφει (header_name, value_as_str) για τη στήλη 'letter'.
    row_index_zero_based: 0 = 1η γραμμή data (όχι headers).
    """
    col1 = col_from_letter(letter)
    if not col1:
        return "", ""
    col0 = col1 - 1
    if col0 < 0 or col0 >= len(df.columns):
        return "", ""
    header = str(df.columns[col0])
    try:
        v = df.iloc[row_index_zero_based, col0]
        return header, "" if pd.isna(v) else str(v)
    except Exception:
        return header, ""

def read_data(xls, sheet_name: str) -> Optional[pd.DataFrame]:
    """Δέχεται .xlsx ή .csv (auto-detect). Επιστρέφει DataFrame ή None."""
    try:
        fname = getattr(xls, "name", "")
        if fname.lower().endswith(".csv"):
            st.write("📑 Sheets:", ["CSV Data"])
            return pd.read_csv(xls)
        # default: xlsx
        xfile = pd.ExcelFile(xls, engine="openpyxl")
        st.write("📑 Sheets:", xfile.sheet_names)
        if sheet_name not in xfile.sheet_names:
            st.error(f"Το sheet '{sheet_name}' δεν βρέθηκε. Διαθέσιμα: {xfile.sheet_names}")
            return None
        return pd.read_excel(xfile, sheet_name=sheet_name, engine="openpyxl")
    except Exception as e:
        st.error(f"Δεν άνοιξε το αρχείο: {e}")
        return None

def safe(v):
    return "" if (v is None or (isinstance(v, float) and pd.isna(v))) else v

# ───────────────────────────── SIDEBAR ─────────────────────────────
with st.sidebar:
    st.header("Ρυθμίσεις")
    debug_mode = st.toggle("🛠 Debug mode", value=True)
    test_mode  = st.toggle("🧪 Test mode (πρώτες 50 γραμμές)", value=True)

    st.subheader("Templates (.docx)")
    tpl_bex    = st.file_uploader("BEX template", type=["docx"], key="tpl_bex")
    tpl_nonbex = st.file_uploader("Non-BEX template", type=["docx"], key="tpl_nonbex")

    st.caption("Placeholders διαθέσιμα στα .docx: "
               "[[title]] [[plan_month]] [[store]] [[bex]] [[plan_vs_target]] "
               "[[mobile_actual]] [[mobile_target]] [[fixed_target]] [[fixed_actual]] "
               "[[voice_vs_target]] [[fixed_vs_target]] [[llu_actual]] [[nga_actual]] "
               "[[ftth_actual]] [[eon_tv_actual]] [[fwa_actual]] "
               "[[mobile_upgrades]] [[fixed_upgrades]] [[pending_mobile]] [[pending_fixed]]")

# ───────────────────────────── ΚΥΡΙΟ INPUT ─────────────────────────────
st.markdown("### 1) Ανέβασε Excel/CSV")
xls = st.file_uploader("Excel/CSV", type=["xlsx", "csv"], key="xls")
sheet_name = st.text_input("Όνομα φύλλου (Sheet - μόνο για Excel)", value="Sheet1")

# ───────────────────────────── STORE & BEX ─────────────────────────────
with st.expander("🏷️ STORE & BEX"):
    st.write("Δώσε στήλη STORE (aliases ή manual γράμμα) και πηγή BEX.")
    # STORE από headers (aliases) ή manual γράμμα
    store_letter = st.text_input("Γράμμα στήλης για STORE (προαιρετικό — αν μείνει κενό θα προσπαθήσει από header)",
                                 value="", placeholder="π.χ. D ή AA")
    bex_mode = st.radio("Πώς βρίσκουμε αν είναι BEX;", ["Από λίστα stores", "Από στήλη (YES/NO)"], index=0)
    manual_bex_list = st.text_input("Λίστα BEX stores (comma-separated)",
                                    "DRZ01, FKM01, ESC01, LND01, PKK01")
    bex_yesno_letter = ""
    if bex_mode == "Από στήλη (YES/NO)":
        bex_yesno_letter = st.text_input("Γράμμα στήλης BEX (YES/NO)", value="", placeholder="π.χ. J")

# ───────────────────────────── MANUAL MAPPING (γράμματα) ─────────────────────────────
with st.expander("🔠 Mapping με γράμματα Excel (A, N, AA, AB, AF, AH)"):
    letter_plan_vs      = st.text_input("plan_vs_target", value="A")
    letter_mobile_tgt   = st.text_input("mobile_target", value="O")
    letter_fixed_tgt    = st.text_input("fixed_target", value="P")
    letter_fixed_act    = st.text_input("fixed_actual", value="Q")
    letter_voice_vs     = st.text_input("voice_vs_target", value="R")
    letter_fixed_vs     = st.text_input("fixed_vs_target", value="S")
    letter_mobile_act   = st.text_input("mobile_actual", value="N")
    letter_llu          = st.text_input("llu_actual", value="T")
    letter_nga          = st.text_input("nga_actual", value="U")
    letter_ftth         = st.text_input("ftth_actual", value="V")
    letter_eon          = st.text_input("eon_tv_actual", value="X")
    letter_fwa          = st.text_input("fwa_actual", value="Y")
    letter_mob_upg      = st.text_input("mobile_upgrades", value="AA")
    letter_fix_upg      = st.text_input("fixed_upgrades", value="AB")
    letter_pending_mob  = st.text_input("pending_mobile", value="AF")
    letter_pending_fix  = st.text_input("pending_fixed", value="AH")

# ───────────────────────────── RUN ─────────────────────────────
run = st.button("🔧 Generate")

if run:
    # Αρχικοί έλεγχοι
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

    # Διαβάζουμε δεδομένα
    df = read_data(xls, sheet_name)
    if df is None or df.empty:
        st.error("Δεν βρέθηκαν δεδομένα στο αρχείο.")
        st.stop()

    st.success(f"OK: {len(df)} γραμμές, {len(df.columns)} στήλες.")
    if debug_mode:
        st.write("**Headers όπως τους βλέπουμε:**", list(df.columns))
        st.dataframe(df.head(10))

    # Πόθεν παίρνουμε STORE:
    # Αν store_letter δόθηκε, θα το διαβάσουμε ως γράμμα· αλλιώς ψάχνουμε aliases.
    store_aliases = ["Dealer_Code", "Dealer code", "Shop Code", "Shop_Code", "ShopCode",
                     "Shop code", "STORE", "Κατάστημα", "store", "dealer_code"]

    # Προview mapping (από 2η γραμμή data)
    with st.expander("🔎 Mapping preview (από 2η γραμμή)"):
        st.write("**Headers:**", list(df.columns))
        sample_row_idx = 1 if len(df) > 1 else 0

        # STORE
        if store_letter.strip():
            store_header, store_value = get_cell_by_letter(df, store_letter, sample_row_idx)
        else:
            chosen = pick(df.columns, *store_aliases)
            store_header = chosen
            store_value = "" if not chosen else ("" if pd.isna(df.iloc[sample_row_idx][chosen]) else str(df.iloc[sample_row_idx][chosen]))

        preview = {
            "store_letter": {"letter": store_letter or "(auto header)", "header": store_header, "sample_row2": store_value},
            "plan_vs_target": {"letter": letter_plan_vs, "sample_row2": get_cell_by_letter(df, letter_plan_vs, sample_row_idx)[1]},
            "mobile_actual": {"letter": letter_mobile_act, "sample_row2": get_cell_by_letter(df, letter_mobile_act, sample_row_idx)[1]},
            "mobile_target": {"letter": letter_mobile_tgt, "sample_row2": get_cell_by_letter(df, letter_mobile_tgt, sample_row_idx)[1]},
            "fixed_target": {"letter": letter_fixed_tgt, "sample_row2": get_cell_by_letter(df, letter_fixed_tgt, sample_row_idx)[1]},
            "fixed_actual": {"letter": letter_fixed_act, "sample_row2": get_cell_by_letter(df, letter_fixed_act, sample_row_idx)[1]},
            "voice_vs_target": {"letter": letter_voice_vs, "sample_row2": get_cell_by_letter(df, letter_voice_vs, sample_row_idx)[1]},
            "fixed_vs_target": {"letter": letter_fixed_vs, "sample_row2": get_cell_by_letter(df, letter_fixed_vs, sample_row_idx)[1]},
            "llu_actual": {"letter": letter_llu, "sample_row2": get_cell_by_letter(df, letter_llu, sample_row_idx)[1]},
            "nga_actual": {"letter": letter_nga, "sample_row2": get_cell_by_letter(df, letter_nga, sample_row_idx)[1]},
            "ftth_actual": {"letter": letter_ftth, "sample_row2": get_cell_by_letter(df, letter_ftth, sample_row_idx)[1]},
            "eon_tv_actual": {"letter": letter_eon, "sample_row2": get_cell_by_letter(df, letter_eon, sample_row_idx)[1]},
            "fwa_actual": {"letter": letter_fwa, "sample_row2": get_cell_by_letter(df, letter_fwa, sample_row_idx)[1]},
            "mobile_upgrades": {"letter": letter_mob_upg, "sample_row2": get_cell_by_letter(df, letter_mob_upg, sample_row_idx)[1]},
            "fixed_upgrades": {"letter": letter_fix_upg, "sample_row2": get_cell_by_letter(df, letter_fix_upg, sample_row_idx)[1]},
            "pending_mobile": {"letter": letter_pending_mob, "sample_row2": get_cell_by_letter(df, letter_pending_mob, sample_row_idx)[1]},
            "pending_fixed": {"letter": letter_pending_fix, "sample_row2": get_cell_by_letter(df, letter_pending_fix, sample_row_idx)[1]},
        }
        st.json(preview, expanded=False)

    # Προετοιμασία templates
    tpl_bex_bytes    = tpl_bex.read()
    tpl_nonbex_bytes = tpl_nonbex.read()

    out_zip = io.BytesIO()
    zf = zipfile.ZipFile(out_zip, "w", zipfile.ZIP_DEFLATED)
    built = 0

    # BEX λίστα
    bex_set = set(s.strip().upper() for s in manual_bex_list.split(",") if s.strip())

    pbar = st.progress(0, text="Δημιουργία εγγράφων…")
    total_rows = len(df) if not test_mode else min(50, len(df))

    # Παίρνουμε index της στήλης STORE με βάση letter ή aliases
    def get_store_for_row(row_idx: int) -> str:
        if store_letter.strip():
            _, v = get_cell_by_letter(df, store_letter, row_idx)
            return (v or "").strip().upper()
        chosen = pick(df.columns, *store_aliases)
        if not chosen:
            return ""
        v = df.iloc[row_idx][chosen]
        return "" if pd.isna(v) else str(v).strip().upper()

    def val(letter: Optional[str], row_idx: int):
        return get_cell_by_letter(df, letter, row_idx)[1]

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
                bex_text_raw = val(bex_yesno_letter, i).strip().lower()
                is_bex = bex_text_raw in ("yes", "y", "1", "true", "ναι")
                bex_text = "YES" if is_bex else "NO"

            # Φτιάχνουμε mapping για placeholders
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
                "fixed_upgrades":   safe(val(letter_fix_upg, i)),
                "pending_mobile":   safe(val(letter_pending_mob, i)),
                "pending_fixed":    safe(val(letter_pending_fix, i)),
            }

            # Γεμίζουμε template
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
        st.error("Δεν δημιουργήθηκε αρχείο. Έλεγξε STORE mapping & templates.")
    else:
        st.success(f"Έτοιμα {built} αρχεία.")
        st.download_button("⬇️ Κατέβασε ZIP", data=out_zip.getvalue(), file_name="reviews_from_excel.zip")