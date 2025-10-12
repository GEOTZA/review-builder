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

# ───────────────────────────── UI CONFIG ─────────────────────────────
st.set_page_config(page_title="Excel → Review/Plan Generator", layout="wide")
st.title("📊 Excel/CSV → 📄 Review/Plan Generator (BEX & Non-BEX)")

# ───────────────────────────── HELPERS ─────────────────────────────
def letter_to_index(letter: str) -> int:
    s = str(letter).strip().upper()
    if not s:
        raise ValueError("Empty letter")
    n = 0
    for ch in s:
        if not ('A' <= ch <= 'Z'):
            raise ValueError(f"Bad column letter: {letter}")
        n = n * 26 + (ord(ch) - ord('A') + 1)
    return n - 1

def excel_letter_to_colname(df: pd.DataFrame, letter: str) -> str | None:
    if not letter or not letter.strip():
        return None
    idx = letter_to_index(letter)
    if idx < 0 or idx >= len(df.columns):
        return None
    return str(df.columns[idx])
def set_default_font(doc: Document, font_name: str = "Aptos") -> None:
    """Ορίζει προεπιλεγμένη γραμματοσειρά σε styles (και eastAsia/complex)."""
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

def cell(row: pd.Series, col: Optional[str]):
    if not col:
        return ""
    v = row[col]
    return "" if pd.isna(v) else v

# --- Excel letters -> 0-based index (A=0, ..., Z=25, AA=26, ...) ---
def xlcol_to_idx(col_letter: str) -> int:
    s = str(col_letter).strip().upper()
    if not s:
        return -1
    n = 0
    for ch in s:
        if not ('A' <= ch <= 'Z'):
            return -1
        n = n * 26 + (ord(ch) - ord('A') + 1)
    return n - 1

def col_from_letter(df: pd.DataFrame, letter: Optional[str]) -> str:
    if not letter:
        return ""
    i = xlcol_to_idx(letter)
    if i < 0 or i >= len(df.columns):
        return ""
    return df.columns[i]

def resolve_col(df: pd.DataFrame, auto_name: str, letter: Optional[str]) -> str:
    return (col_from_letter(df, letter) if letter else (auto_name or "")) or ""

def truthy(val) -> bool:
    s = str(val).strip().lower()
    return s in {"yes", "y", "true", "1", "ναι", "nai", "✔", "✓"}

def fmt_percent(val):
    """Δέχεται 0.85 ή 85 ή '85%' -> '85%'"""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    s = str(val).strip()
    if not s:
        return ""
    if s.endswith("%"):
        try:
            return f"{float(s[:-1].replace(',', '.')):.0f}%"
        except Exception:
            return s
    try:
        num = float(s.replace(",", "."))
        if num <= 1:
            num *= 100
        return f"{num:.0f}%"
    except Exception:
        return s

def read_data(xls, sheet_name: str) -> Optional[pd.DataFrame]:
    """Δέχεται .xlsx ή .csv (auto-detect από το όνομα). Επιστρέφει DataFrame ή None."""
    try:
        fname = getattr(xls, "name", "")
        if fname.lower().endswith(".csv"):
            st.write("📑 Sheets:", ["CSV Data"])
            return pd.read_csv(xls)
        # default: xlsx
        xls.seek(0)
        xfile = pd.ExcelFile(xls, engine="openpyxl")
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
test_mode  = st.sidebar.toggle("🧪 Test mode (limit rows=50)", value=True)

st.sidebar.header("⚙️ BEX")
bex_mode = st.sidebar.radio("Πηγή BEX", ["Στήλη στο Excel", "Λίστα (comma-separated)"], index=1)
bex_default = "DRZ01,FKM01,ESC01,LND01,PKK01"
bex_list = set()
if bex_mode == "Λίστα (comma-separated)":
    bex_input = st.sidebar.text_area("BEX stores", bex_default)
    bex_list = set(s.strip().upper() for s in bex_input.split(",") if s.strip())

st.sidebar.subheader("📄 Templates (.docx)")
tpl_bex    = st.sidebar.file_uploader("BEX template", type=["docx"])
tpl_nonbex = st.sidebar.file_uploader("Non-BEX template", type=["docx"])
st.sidebar.caption(
    "Placeholders: [[title]], [[store]], [[plan_month]], [[plan_vs_target]], [[mobile_actual]], [[mobile_target]], "
    "[[fixed_actual]], [[fixed_target]], [[voice_vs_target]], [[fixed_vs_target]], [[llu_actual]], [[nga_actual]], "
    "[[ftth_actual]], [[eon_tv_actual]], [[fwa_actual]], [[mobile_upgrades]], [[fixed_upgrades]], "
    "[[pending_mobile]], [[pending_fixed]]"
)

st.sidebar.subheader("📌 Manual mapping (Excel letters)")
# Από το mapping που έδωσες (γράμματα/διγράμματα). Αλλάζουν από το UI.
def resolve_letters_preview(df: pd.DataFrame, mapping_letters: dict[str, str]) -> dict[str, str | None]:
    out = {}
    for k, L in mapping_letters.items():
        out[k] = excel_letter_to_colname(df, L) if L and L.strip() else None
    return out

letters_map = {
    "plan_vs_target": letter_plan_vs,
    "mobile_plan": letter_mobile_plan,
    "mobile_actual": letter_mobile_act,
    "mobile_target": letter_mobile_tgt,
    "fixed_target": letter_fixed_tgt,
    "fixed_actual": letter_fixed_act,
    "voice_vs_target": letter_voice_vs,
    "fixed_vs_target": letter_fixed_vs,
    "llu_actual": letter_llu,
    "nga_actual": letter_nga,
    "ftth_actual": letter_ftth,
    "eon_tv_actual": letter_eon,
    "fwa_actual": letter_fwa,
    "mobile_upgrades": letter_mob_upg,
    "fixed_upgrades": letter_fix_upg,
    "pending_mobile": letter_pend_mob,
    "pending_fixed": letter_pend_fix,
}

st.markdown("#### 🧭 Letters → Headers (live)")
if xls:
    _dfp = read_data(xls, sheet_name)
    if _dfp is not None and not _dfp.empty:
        st.json(resolve_letters_preview(_dfp, letters_map))
        st.caption("Αν κάποιο key δείχνει σε λάθος header (π.χ. 'Dealer_Code'), άλλαξε το γράμμα ή το Sheet.")
L_PLAN_VS   = st.sidebar.text_input("plan vs target", value="A")
L_MOB_PLAN  = st.sidebar.text_input("mobile plan (optional)", value="B")
L_BEXCOL    = st.sidebar.text_input("BEX (YES/NO) column", value="J")

L_MOB_ACT   = st.sidebar.text_input("MOBILE ACTUAL", value="N")
L_MOB_TGT   = st.sidebar.text_input("mobile target", value="O")
L_FIX_TGT   = st.sidebar.text_input("fixed target", value="P")
L_FIX_ACT   = st.sidebar.text_input("total fixed actual", value="Q")

L_VOICE_VS  = st.sidebar.text_input("voice Vs target %", value="R")
L_FIXED_VS  = st.sidebar.text_input("fixed vs target %", value="S")

L_LLU       = st.sidebar.text_input("llu actual", value="T")
L_NGA       = st.sidebar.text_input("nga actual", value="U")
L_FTTH      = st.sidebar.text_input("ftth actual", value="V")
L_EON       = st.sidebar.text_input("eon tv actual", value="X")
L_FWA       = st.sidebar.text_input("fwa actual", value="Y")

L_MOB_UPG   = st.sidebar.text_input("mobile upgrades", value="aa")
L_FIX_UPG   = st.sidebar.text_input("fixed upgrades", value="ab")
L_PEND_MOB  = st.sidebar.text_input("total pending mobile", value="af")
L_PEND_FIX  = st.sidebar.text_input("total pending fixed", value="ah")

L_STORE     = st.sidebar.text_input("STORE (Shop Code) override", value="")
st.sidebar.caption("Αν αφήσεις κενό, θα γίνει auto-map από headers (Shop Code, STORE κ.λπ.).")

st.sidebar.subheader("👀 Preview letters")
preview_letters = st.sidebar.text_input(
    "π.χ. A,B,J,N,O,P,Q,R,S,T,U,V,X,Y,aa,ab,af,ah",
    value="A,J,N,O,P,Q,R,S,V,Y,aa,ab,af,ah"
)

# ───────────────────────────── MAIN INPUTS ─────────────────────────────
st.markdown("### 1) Ανέβασε Excel/CSV")
xls = st.file_uploader("Excel/CSV", type=["xlsx", "csv"])
sheet_name = st.text_input("Όνομα φύλλου (Sheet - μόνο για Excel)", value="Sheet1")

run = st.button("🔧 Generate")

# ───────────────────────────── MAIN ─────────────────────────────
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
        st.write("🔍 Headers:", list(df.columns))
        st.dataframe(df.head(10))

    cols = list(df.columns)

    # AUTO-MAP από headers
    col_store   = pick(cols, "Shop Code","Shop_Code","ShopCode","Shop code","STORE","Κατάστημα", r"shop.?code","dealer code")
    col_bex     = pick(cols, "BEX store","BEX","is_bex","bex_flag", r"\bbex\b", r"bex.*store", r"is.*bex")
    col_plan_vs = pick(cols, "plan vs target","plan_vs_target","% plan", r"plan.*vs.*target")

    col_mob_act = pick(cols, "MOBILE ACTUAL","mobile actual", r"mobile.*actual")
    col_mob_tgt = pick(cols, "mobile target", r"mobile.*target","mobile plan")
    col_fix_tgt = pick(cols, "fixed target","target fixed", r"fixed.*target","fixed plan total","fixed plan")
    col_fix_act = pick(cols, "total fixed actual","total fixed", r"(total|sum).?fixed.*actual","fixed actual")

    col_voice_vs= pick(cols, "voice Vs target", r"voice.*vs.*target")
    col_fixed_vs= pick(cols, "fixed vs target", r"fixed.*vs.*target")

    col_llu     = pick(cols, "llu actual", r"llu.*actual")
    col_nga     = pick(cols, "nga actual", r"nga.*actual")
    col_ftth    = pick(cols, "ftth actual", r"ftth.*actual")
    col_eon     = pick(cols, "eon tv actual", r"(eon|tv).*actual")
    col_fwa     = pick(cols, "fwa actual", r"fwa.*actual")

    col_mob_upg = pick(cols, "mobile upgrades", r"mobile.*upg")
    col_fix_upg = pick(cols, "fixed upgrades", r"fixed.*upg")
    col_pend_m  = pick(cols, "total pending mobile", r"pending.*mobile")
    col_pend_f  = pick(cols, "total pending fixed", r"pending.*fixed")

    # OVERRIDE με γράμματα (αν έχουν δοθεί)
    col_store   = resolve_col(df, col_store,   L_STORE)
    col_bex     = resolve_col(df, col_bex,     L_BEXCOL)
    col_plan_vs = resolve_col(df, col_plan_vs, L_PLAN_VS)

    col_mob_act = resolve_col(df, col_mob_act, L_MOB_ACT)
    col_mob_tgt = resolve_col(df, col_mob_tgt, L_MOB_TGT)
    col_fix_tgt = resolve_col(df, col_fix_tgt, L_FIX_TGT)
    col_fix_act = resolve_col(df, col_fix_act, L_FIX_ACT)

    col_voice_vs= resolve_col(df, col_voice_vs, L_VOICE_VS)
    col_fixed_vs= resolve_col(df, col_fixed_vs, L_FIXED_VS)

    col_llu     = resolve_col(df, col_llu,     L_LLU)
    col_nga     = resolve_col(df, col_nga,     L_NGA)
    col_ftth    = resolve_col(df, col_ftth,    L_FTTH)
    col_eon     = resolve_col(df, col_eon,     L_EON)
    col_fwa     = resolve_col(df, col_fwa,     L_FWA)

    col_mob_upg = resolve_col(df, col_mob_upg, L_MOB_UPG)
    col_fix_upg = resolve_col(df, col_fix_upg, L_FIX_UPG)
    col_pend_m  = resolve_col(df, col_pend_m,  L_PEND_MOB)
    col_pend_f  = resolve_col(df, col_pend_f,  L_PEND_FIX)

    # Εμφάνιση mapping
    with st.expander("Χαρτογράφηση (auto + manual)"):
        st.write({
            "STORE": col_store, "BEX": col_bex,
            "plan_vs_target": col_plan_vs,
            "mobile_actual": col_mob_act, "mobile_target": col_mob_tgt,
            "fixed_target": col_fix_tgt, "fixed_actual": col_fix_act,
            "voice_vs_target": col_voice_vs, "fixed_vs_target": col_fixed_vs,
            "llu_actual": col_llu, "nga_actual": col_nga, "ftth_actual": col_ftth,
            "eon_tv_actual": col_eon, "fwa_actual": col_fwa,
            "mobile_upgrades": col_mob_upg, "fixed_upgrades": col_fix_upg,
            "pending_mobile": col_pend_m, "pending_fixed": col_pend_f
        })

    if not col_store:
        st.error("Δεν βρέθηκε στήλη STORE (π.χ. 'Shop Code'). Διόρθωσε την κεφαλίδα ή δώσε γράμμα στήλης.")
        st.stop()

    # Προεπισκόπηση με γράμματα
    if preview_letters.strip():
        letters = [s.strip() for s in preview_letters.split(",") if s.strip()]
        preview_cols = [c for c in (col_from_letter(df, s) for s in letters) if c]
        if preview_cols:
            st.write("🔎 **Preview**:", preview_cols)
            st.dataframe(df[preview_cols].head(20))

    # Διαβάζουμε μοναδικό κελί B17 για "μήνας Πλάνου"
    plan_month = ""
    try:
        from openpyxl import load_workbook
        xls.seek(0)
        wb = load_workbook(filename=xls, read_only=True, data_only=True)
        ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.active
        plan_month = str(ws["B17"].value or "").strip()
    except Exception:
        plan_month = ""

    # Templates → bytes
    tpl_bex_bytes    = tpl_bex.read()
    tpl_nonbex_bytes = tpl_nonbex.read()

    # Out ZIP
    out_zip = io.BytesIO()
    zf = zipfile.ZipFile(out_zip, "w", zipfile.ZIP_DEFLATED)
    built = 0

    pbar = st.progress(0, text="Δημιουργία εγγράφων…")
    total = len(df) if not test_mode else min(50, len(df))
store_col = col_store or excel_letter_to_colname(df, letter_store)  # αν έχεις letter για STORE
conflicts = []
for k, L in letters_map.items():
    hdr = excel_letter_to_colname(df, L) if L and L.strip() else None
    if hdr and store_col and hdr == store_col:
        conflicts.append((k, L, hdr))
if conflicts:
    st.warning(f"⚠️ Κάποια πεδία πέφτουν στη στήλη STORE ({store_col}): {conflicts}")
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
                bex_val = cell(row, col_bex) if col_bex else ""
                is_bex = truthy(bex_val)

            mapping = {
                "title": f"Review September 2025 — Plan October 2025 — {store_up}",
                "store": store_up,
                "plan_month": plan_month,

                "plan_vs_target": fmt_percent(cell(row, col_plan_vs)),
                "mobile_actual":  cell(row, col_mob_act),
                "mobile_target":  cell(row, col_mob_tgt),
                "fixed_actual":   cell(row, col_fix_act),
                "fixed_target":   cell(row, col_fix_tgt),

                "voice_vs_target": fmt_percent(cell(row, col_voice_vs)),
                "fixed_vs_target": fmt_percent(cell(row, col_fixed_vs)),

                "llu_actual":  cell(row, col_llu),
                "nga_actual":  cell(row, col_nga),
                "ftth_actual": cell(row, col_ftth),
                "eon_tv_actual": cell(row, col_eon),
                "fwa_actual":  cell(row, col_fwa),

                "mobile_upgrades": cell(row, col_mob_upg),
                "fixed_upgrades":  cell(row, col_fix_upg),

                "pending_mobile": cell(row, col_pend_m),
                "pending_fixed":  cell(row, col_pend_f),
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