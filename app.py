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
def set_default_font(doc: Document, font_name: str = "Aptos") -> None:
    """Ορίζει προεπιλεγμένη γραμματοσειρά σε όλα τα styles (και eastAsia)."""
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
        try:
            pat = re.compile(a, re.IGNORECASE)
        except re.error:
            continue
        for c in columns:
            if re.search(pat, str(c)):
                return c
    return ""

def letter_to_index(letter: str) -> Optional[int]:
    """
    Μετατρέπει γράμμα Excel σε 0-based index (π.χ. A->0, N->13, AA->26).
    Επιστρέφει None αν είναι κενό.
    """
    s = str(letter or "").strip().upper()
    if not s:
        return None
    # Επιτρέπουμε και "B17" ως αναφορά: παίρνουμε μόνο τα γράμματα
    s = re.sub(r"[^A-Z]", "", s)
    if not s:
        return None
    idx = 0
    for ch in s:
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx - 1

def coerce_number(val) -> Optional[float]:
    """Μετατρέπει σε float αν γίνεται, αλλιώς None."""
    if val is None:
        return None
    if isinstance(val, (int, float)) and pd.notna(val):
        return float(val)
    try:
        s = str(val).strip().replace("%", "")
        if s == "":
            return None
        return float(s)
    except Exception:
        return None

def as_percent(val) -> str:
    """1.22 -> '122%' (χωρίς δεκαδικά)."""
    x = coerce_number(val)
    if x is None:
        return ""
    # Αν ήδη είναι 0-100, μην το ξαναπολλαπλασιάσεις
    if x <= 1.0:
        x = x * 100.0
    return f"{round(x):d}%"

def read_data(xls, sheet_name: str) -> Optional[pd.DataFrame]:
    """Δέχεται .xlsx ή .csv (auto-detect από το όνομα). Επιστρέφει DataFrame ή None."""
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

def value_from_letter_row(df: pd.DataFrame, row_idx_0_based: int, letter: str) -> Any:
    """Δώσε row index (0-based) + γράμμα στήλης, πάρε τιμή."""
    ci = letter_to_index(letter)
    if ci is None:
        return ""
    if row_idx_0_based < 0 or row_idx_0_based >= len(df):
        return ""
    if ci < 0 or ci >= len(df.columns):
        return ""
    val = df.iat[row_idx_0_based, ci]
    return "" if pd.isna(val) else val

# ───────────────────────────── SIDEBAR ─────────────────────────────
debug_mode = st.sidebar.toggle("🛠 Debug mode", value=True)
test_mode  = st.sidebar.toggle("🧪 Test mode (πρώτες 50 γραμμές)", value=False)

st.sidebar.header("📄 Templates (.docx)")
tpl_bex    = st.sidebar.file_uploader("BEX template", type=["docx"])
tpl_nonbex = st.sidebar.file_uploader("Non-BEX template", type=["docx"])
st.sidebar.caption(
    "Placeholders: [[title]], [[plan_month]], [[store]], [[bex]], [[plan_vs_target]], [[mobile_actual]], [[mobile_target]], "
    "[[fixed_target]], [[fixed_actual]], [[voice_vs_target]], [[fixed_vs_target]], [[llu_actual]], [[nga_actual]], [[ftth_actual]], "
    "[[eon_tv_actual]], [[fwa_actual]], [[mobile_upgrades]], [[fixed_upgrades]], [[pending_mobile]], [[pending_fixed]]"
)

st.sidebar.header("🏪 STORE & BEX")
default_bex = "DRZ01,FKM01,ESC01,LND01,PKK01"
bex_mode = st.sidebar.radio("Πώς βρίσκουμε αν είναι BEX:", ["Από λίστα (DRZ01, ...)", "Από στήλη (YES/NO)"], index=0)
bex_list_text = st.sidebar.text_input("BEX stores (comma)", value=default_bex)
bex_yesno_header_hint = st.sidebar.text_input("Όνομα στήλης (YES/NO)", value="BEX store")

# ───────────────────────────── MAIN INPUTS ─────────────────────────────
st.markdown("### 1) Ανέβασε Excel/CSV")
xls = st.file_uploader("Excel/CSV", type=["xlsx", "csv"])
sheet_name = st.text_input("Όνομα φύλλου (Sheet - μόνο για Excel)", value="Sheet1")

with st.expander("📌 Ρύθμιση γραμμών (headers & δεδομένα)"):
    header_row_1based = st.number_input("Header row (1-based)", min_value=1, value=1, step=1,
                                        help="Σε ποια γραμμή βρίσκονται οι κεφαλίδες. Συνήθως 1.")
    data_start_row_1based = st.number_input("Δεδομένα ξεκινούν στη γραμμή (1-based)", min_value=2, value=2, step=1,
                                            help="Η πρώτη γραμμή δεδομένων (συνήθως 2).")

with st.expander("📌 STORE (στήλη ή γράμμα)"):
    store_mode = st.radio("Πηγή Store code:", ["Από κεφαλίδα στήλης", "Με γράμμα Excel"], index=0)
    store_header_fallback = "Dealer_Code"
    store_header_input = st.text_input("Όνομα κεφαλίδας για Store", value=store_header_fallback)
    store_letter = st.text_input("Γράμμα Excel για Store (π.χ. A, G, AA)", value="")

with st.expander("📌 Mapping με γράμματα Excel (A, N, AA, AB, AF, AH)"):
    # Τα γράμματα αυτά είναι **προαιρετικά**. Αν αφεθούν κενά, θα γίνει auto-map από headers.
    letter_plan_vs   = st.text_input("plan_vs_target", value="A")
    letter_mob_act   = st.text_input("mobile_actual", value="N")
    letter_mob_tgt   = st.text_input("mobile_target", value="O")
    letter_fix_tgt   = st.text_input("fixed_target", value="P")
    letter_fix_act   = st.text_input("fixed_actual", value="Q")
    letter_voice_vs  = st.text_input("voice_vs_target (ποσοστό)", value="R")
    letter_fixed_vs  = st.text_input("fixed_vs_target (ποσοστό)", value="S")
    letter_llu       = st.text_input("llu_actual", value="T")
    letter_nga       = st.text_input("nga_actual", value="U")
    letter_ftth      = st.text_input("ftth_actual", value="V")
    letter_eon       = st.text_input("eon_tv_actual", value="X")
    letter_fwa       = st.text_input("fwa_actual", value="Y")
    letter_mob_upg   = st.text_input("mobile_upgrades", value="AA")
    letter_fix_upg   = st.text_input("fixed_upgrades", value="AB")
    letter_pend_mob  = st.text_input("pending_mobile", value="AF")
    letter_pend_fix  = st.text_input("pending_fixed", value="AH")

plan_month_text = st.text_input("Κείμενο για [[plan_month]]", value="Review September 2025 — Plan October 2025")

run = st.button("🔧 Generate")

# ───────────────────────────── MAIN ─────────────────────────────
if run:
    # ── Έλεγχοι αρχείων
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

    # ── Διαβάζουμε δεδομένα
    df_raw = read_data(xls, sheet_name)
    if df_raw is None or df_raw.empty:
        st.error("Δεν βρέθηκαν δεδομένα στο αρχείο.")
        st.stop()

    # Μετακινούμε headers αν ο χρήστης όρισε διαφορετική γραμμή κεφαλίδων
    if header_row_1based != 1:
        new_header = df_raw.iloc[header_row_1based - 1].tolist()
        df = df_raw.iloc[header_row_1based:].copy()
        df.columns = new_header
        df.reset_index(drop=True, inplace=True)
    else:
        df = df_raw.copy()

    st.success(f"OK: {len(df)} γραμμές, {len(df.columns)} στήλες.")
    if debug_mode:
        st.write("Headers όπως βλέπουμε:", list(df.columns))
        st.dataframe(df.head(10))

    # ── Store column resolve
    cols = list(df.columns)
    col_store_auto = pick(
        cols,
        "Dealer Code", "Dealer_Code", "dealer code", "dealer_code",
        "Shop Code", "Shop_Code", "Shop code",
        "STORE", "Κατάστημα", r"shop.?code", r"dealer.?code"
    )
    if store_mode == "Από κεφαλίδα στήλης":
        col_store = store_header_input if store_header_input in cols else col_store_auto
        if not col_store:
            st.error("Δεν εντοπίστηκε στήλη Store. Βάλε σωστή κεφαλίδα ή χρησιμοποίησε γράμμα Excel.")
            st.stop()
    else:
        col_store = ""  # θα διαβάσουμε από γράμμα

    # ── BEX detect
    bex_set = set(s.strip().upper() for s in bex_list_text.split(",") if s.strip())
    col_bex_yesno = bex_yesno_header_hint if bex_yesno_header_hint in cols else pick(cols, "BEX store", "BEX", r"bex.?store")

    # ── Αποθήκευση template bytes
    tpl_bex_bytes    = tpl_bex.read()
    tpl_nonbex_bytes = tpl_nonbex.read()

    # ── Template audit (πρέπει να είναι ΜΕΤΑ τα tpl_*_bytes)
    with st.expander("🔎 Template audit (placeholders που βρέθηκαν στο .docx)"):
        def placeholders_in_doc(doc_bytes: bytes) -> list[str]:
            r = []
            pat = re.compile(r"\[\[([A-Za-z0-9_]+)\]\]")
            doc = Document(io.BytesIO(doc_bytes))
            for p in doc.paragraphs:
                for m in pat.findall(p.text):
                    r.append(m)
            for t in doc.tables:
                for row in t.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            for m in pat.findall(p.text):
                                r.append(m)
            return sorted(set(r))

        st.write("BEX:", placeholders_in_doc(tpl_bex_bytes))
        st.write("Non-BEX:", placeholders_in_doc(tpl_nonbex_bytes))

    # ── Preview mapping (2η γραμμή δεδομένων)
    with st.expander("🔍 Mapping preview (από 2η γραμμή)"):
        row2_idx0 = max(0, data_start_row_1based - 2)
        preview = {
            "store": value_from_letter_row(df, row2_idx0, store_letter) if store_mode == "Με γράμμα Excel"
                     else ("" if not col_store else ("" if pd.isna(df.iloc[row2_idx0][col_store]) else df.iloc[row2_idx0][col_store])),
            "plan_vs_target": value_from_letter_row(df, row2_idx0, letter_plan_vs),
            "mobile_actual":  value_from_letter_row(df, row2_idx0, letter_mob_act),
            "mobile_target":  value_from_letter_row(df, row2_idx0, letter_mob_tgt),
            "fixed_target":   value_from_letter_row(df, row2_idx0, letter_fix_tgt),
            "fixed_actual":   value_from_letter_row(df, row2_idx0, letter_fix_act),
            "voice_vs_target": value_from_letter_row(df, row2_idx0, letter_voice_vs),
            "fixed_vs_target": value_from_letter_row(df, row2_idx0, letter_fixed_vs),
            "llu_actual":     value_from_letter_row(df, row2_idx0, letter_llu),
            "nga_actual":     value_from_letter_row(df, row2_idx0, letter_nga),
            "ftth_actual":    value_from_letter_row(df, row2_idx0, letter_ftth),
            "eon_tv_actual":  value_from_letter_row(df, row2_idx0, letter_eon),
            "fwa_actual":     value_from_letter_row(df, row2_idx0, letter_fwa),
            "mobile_upgrades": value_from_letter_row(df, row2_idx0, letter_mob_upg),
            "fixed_upgrades":  value_from_letter_row(df, row2_idx0, letter_fix_upg),
            "pending_mobile":  value_from_letter_row(df, row2_idx0, letter_pend_mob),
            "pending_fixed":   value_from_letter_row(df, row2_idx0, letter_pend_fix),
        }
        st.write(preview)

    # ── Έξοδος ZIP
    out_zip = io.BytesIO()
    zf = zipfile.ZipFile(out_zip, "w", zipfile.ZIP_DEFLATED)
    built = 0

    pbar = st.progress(0, text="Δημιουργία εγγράφων…")
    total_rows = len(df)
    if test_mode:
        total_rows = min(total_rows, 50)

    # ── Loop γραμμών
    for i0, row in enumerate(df.itertuples(index=False), start=0):
        # skip πριν από data_start_row
        if i0 < (data_start_row_1based - 1):
            continue
        if test_mode and (i0 - (data_start_row_1based - 1)) >= 50:
            st.info("🧪 Test mode: σταμάτησα στις 50 γραμμές.")
            break

        idx1 = i0 + 1  # 1-based για UI
        try:
            # Ανάγνωση store
            if store_mode == "Με γράμμα Excel":
                store_val = value_from_letter_row(df, i0, store_letter)
            else:
                store_val = "" if not col_store else getattr(row, col_store, "")
            store = "" if pd.isna(store_val) else str(store_val).strip().upper()
            if not store:
                pbar.progress(min((i0 + 1) / max(total_rows, 1), 1.0), text=f"Παράλειψη γραμμής {idx1} (κενό store)")
                continue

            # BEX flag
            if bex_mode == "Από λίστα (DRZ01, ...)":
                is_bex = store in bex_set
            else:
                raw = "" if not col_bex_yesno else getattr(row, col_bex_yesno, "")
                raw = "" if pd.isna(raw) else str(raw).strip().lower()
                is_bex = raw in ("yes", "y", "1", "true", "ναι")

            # Λήψη τιμών από γράμματα (πάντα έχουν προτεραιότητα αν δοθούν)
            def pick_val(letter: str, header_fallbacks) -> Any:
                if letter.strip():
                    return value_from_letter_row(df, i0, letter)
                # αλλιώς από headers (auto-map)
                for h in header_fallbacks:
                    h_real = pick(cols, h)
                    if h_real:
                        v = getattr(row, h_real, "")
                        return "" if pd.isna(v) else v
                return ""

            v_plan_vs  = pick_val(letter_plan_vs,  ["plan vs target", r"plan.*vs.*target"])
            v_mob_act  = pick_val(letter_mob_act,  ["mobile actual", r"mobile.*actual", "BNS VOICE"])
            v_mob_tgt  = pick_val(letter_mob_tgt,  ["mobile target", r"mobile.*target", "mobile plan", "target voice"])
            v_fix_tgt  = pick_val(letter_fix_tgt,  ["fixed target", r"fixed.*target", "target fixed"])
            v_fix_act  = pick_val(letter_fix_act,  ["total fixed", r"(total|sum).?fixed.*actual", "fixed actual"])
            v_voice_vs = pick_val(letter_voice_vs, ["% voice", "voice vs target"])
            v_fixed_vs = pick_val(letter_fixed_vs, ["% fixed", "fixed vs target"])
            v_llu      = pick_val(letter_llu,      ["llu actual"])
            v_nga      = pick_val(letter_nga,      ["nga actual"])
            v_ftth     = pick_val(letter_ftth,     ["ftth actual"])
            v_eon      = pick_val(letter_eon,      ["eon tv actual"])
            v_fwa      = pick_val(letter_fwa,      ["fwa actual"])
            v_mupg     = pick_val(letter_mob_upg,  ["mobile upgrades"])
            v_fupg     = pick_val(letter_fix_upg,  ["fixed upgrades"])
            v_pmob     = pick_val(letter_pend_mob, ["total pending mobile"])
            v_pfix     = pick_val(letter_pend_fix, ["total pending fixed"])

            # Μορφοποίηση ποσοστών
            plan_vs_fmt  = as_percent(v_plan_vs) if v_plan_vs != "" else ""
            voice_vs_fmt = as_percent(v_voice_vs) if v_voice_vs != "" else ""
            fixed_vs_fmt = as_percent(v_fixed_vs) if v_fixed_vs != "" else ""

            mapping = {
                "title": f"Review September 2025 — Plan October 2025 — {store}",
                "plan_month": plan_month_text,
                "store": store,
                "bex": "YES" if is_bex else "NO",
                "plan_vs_target": plan_vs_fmt or v_plan_vs,
                "mobile_actual":  v_mob_act,
                "mobile_target":  v_mob_tgt,
                "fixed_target":   v_fix_tgt,
                "fixed_actual":   v_fix_act,
                "voice_vs_target": voice_vs_fmt or v_voice_vs,
                "fixed_vs_target": fixed_vs_fmt or v_fixed_vs,
                "llu_actual":     v_llu,
                "nga_actual":     v_nga,
                "ftth_actual":    v_ftth,
                "eon_tv_actual":  v_eon,
                "fwa_actual":     v_fwa,
                "mobile_upgrades": v_mupg,
                "fixed_upgrades":  v_fupg,
                "pending_mobile":  v_pmob,
                "pending_fixed":   v_pfix,
            }

            # Δημιουργία εγγράφου
            tpl_bytes = tpl_bex_bytes if is_bex else tpl_nonbex_bytes
            doc = Document(io.BytesIO(tpl_bytes))
            set_default_font(doc, "Aptos")
            replace_placeholders(doc, mapping)

            out_name = f"{store}_ReviewSep_PlanOct.docx"
            buf = io.BytesIO()
            doc.save(buf)
            zf.writestr(out_name, buf.getvalue())
            built += 1

            pbar.progress(min((i0 + 1) / max(len(df), 1), 1.0), text=f"Φτιάχνω: {out_name} ({i0 + 1}/{len(df)})")

        except Exception as e:
            st.warning(f"⚠️ Σφάλμα στη γραμμή Excel {i0 + 1}: {e}")
            if debug_mode:
                st.exception(e)

    zf.close()
    pbar.empty()

    if built == 0:
        st.error("Δεν δημιουργήθηκε αρχείο. Έλεγξε STORE mapping, letters & templates.")
    else:
        st.success(f"Έτοιμα {built} αρχεία.")
        st.download_button("⬇️ Κατέβασε ZIP", data=out_zip.getvalue(), file_name="reviews_from_excel.zip")