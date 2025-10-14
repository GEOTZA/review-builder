# app.py — Nova Letters / Review Builder (robust)
import io
import re
import json
import zipfile
import unicodedata
import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, Tuple

import streamlit as st
from docx import Document
from docx.text.paragraph import Paragraph
from docx.table import _Cell as TableCell

try:
    import pandas as pd  # needs pandas, openpyxl in requirements.txt
except Exception:
    pd = None

# ───────────────────────────── CONFIG ─────────────────────────────
st.set_page_config(page_title="Nova Letters — Batch Builder", layout="wide")
APP_TITLE = "📄 Nova Letters — Μαζική Παραγωγή (BEX / NON-BEX)"

HERE = Path(__file__).parent
RUNTIME = HERE / "runtime"
RUNTIME.mkdir(exist_ok=True)

TEMPLATES_DIR = HERE / "templates"
DEFAULT_TEMPLATE = TEMPLATES_DIR / "default.docx"
REPO_MAPPING = HERE / "store_mapping.json"  # προαιρετικό json για ονόματα/κατηγορίες/προεπιλογή template

# ───────────────────────────── HELPERS ─────────────────────────────
def _norm_header(s: str) -> str:
    """Normalize headers: αφαιρεί τόνους/μη ASCII, κατεβάζει πεζά, αντικαθιστά κενά/σύμβολα με _."""
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode("ascii")
    s = s.strip().lower()
    s = re.sub(r"[^a-z0-9]+", "_", s)
    return s.strip("_")

def _is_nan(x: Any) -> bool:
    try:
        import math
        return x is None or (isinstance(x, float) and math.isnan(x))
    except Exception:
        return x is None

def _safe_str(x: Any) -> str:
    if _is_nan(x):
        return ""
    return str(x)

def format_percent(x: Any) -> str:
    """1.22 -> 122% , 0.87 -> 87% , 87 -> 87%"""
    if _is_nan(x) or x == "":
        return ""
    try:
        val = float(x)
    except Exception:
        return str(x)
    # αν είναι < 1 το θεωρούμε αναλογία (0.87 => 87%)
    if val < 1:
        return f"{val * 100:.0f}%"
    # αν είναι μεταξύ 1..10 (π.χ. 1.22 => 122%)
    if val < 10:
        return f"{val * 100:.0f}%"
    # αλλιώς ήδη είναι % (87 => 87%)
    return f"{val:.0f}%"

def load_store_mapping(path: Path | None) -> Dict[str, Any]:
    if not path or not path.exists():
        return {}
    with path.open("r", encoding="utf-8") as f:
        return json.load(f)

def pick_template_path(
    store_template_name: str | None,
    category: str | None,
    uploaded_template: Path | None,
    tpl_bex: Path | None,
    tpl_nonbex: Path | None,
) -> Path:
    # 1) Global uploaded (override για όλους)
    if uploaded_template and uploaded_template.exists():
        return uploaded_template
    # 2) Category specific
    cat = (category or "NON_BEX").upper()
    if cat == "BEX" and tpl_bex and tpl_bex.exists():
        return tpl_bex
    if cat != "BEX" and tpl_nonbex and tpl_nonbex.exists():
        return tpl_nonbex
    # 3) Per-store template από templates/
    candidate = TEMPLATES_DIR / (store_template_name or "default.docx")
    if candidate.exists():
        return candidate
    # 4) Fallback
    return DEFAULT_TEMPLATE

# ---- Placeholder extraction from docx (για audit) ----
PLACEHOLDER_RE = re.compile(r"\[\[([A-Za-z0-9_]+)\]\]")

def extract_placeholders_from_paragraph(p: Paragraph) -> Iterable[str]:
    # Join full text of paragraph (runs μπορεί να έχουν κόψει τα tokens)
    text = "".join(run.text for run in p.runs)
    return (m.group(1) for m in PLACEHOLDER_RE.finditer(text))

def extract_placeholders_from_doc(doc: Document) -> set[str]:
    found: set[str] = set()
    for p in doc.paragraphs:
        found.update(extract_placeholders_from_paragraph(p))
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    found.update(extract_placeholders_from_paragraph(p))
    return found

# ---- Robust replace across runs ----
def replace_placeholders_in_paragraph(p: Paragraph, mapping: Dict[str, Any]) -> None:
    text = "".join(run.text for run in p.runs)
    if not text:
        return
    # κάνουμε replace σε όλο το paragraph text
    for k, v in mapping.items():
        text = text.replace(f"[[{k}]]", _safe_str(v))
    # καθαρίζουμε runs και ξαναγράφουμε ένα run με το αποτέλεσμα
    for _ in range(len(p.runs) - 1, -1, -1):
        p.runs[_].clear()  # clear text
    # δεν υπάρχει επίσημο API για να "αδειάσεις" σωστά, οπότε:
    p.clear()
    p.add_run(text)

def replace_all(doc: Document, mapping: Dict[str, Any]) -> None:
    for p in doc.paragraphs:
        replace_placeholders_in_paragraph(p, mapping)
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    replace_placeholders_in_paragraph(p, mapping)

# ---- Build mapping για ένα store (Excel row) ----
PERCENT_HINT_SUFFIXES = ("_vs_target", "_pct", "_percent", "_percentage")

def build_placeholder_map(store_code: str, store_name: str, row_dict: Dict[str, Any]) -> Dict[str, Any]:
    today = datetime.date.today()
    out: Dict[str, Any] = {
        "store_code": store_code,
        "store_name": store_name,
        "month_name": today.strftime("%B"),
        "year": today.year,
        # convenience title/plan_month placeholders (αν τα χρειαστείς στα templates)
        "title": f"Review {today.strftime('%B %Y')} — Plan {(today.replace(day=1) + datetime.timedelta(days=32)).strftime('%B %Y')}",
        "plan_month": f"Review {today.strftime('%B %Y')} — Plan {(today.replace(day=1) + datetime.timedelta(days=32)).strftime('%B %Y')}",
    }
    # πέρασε ΟΛΕΣ τις στήλες του excel ως [[normalized_header]]
    for k, v in row_dict.items():
        if any(k.endswith(suf) for suf in PERCENT_HINT_SUFFIXES):
            out[k] = format_percent(v)
        else:
            out[k] = "" if _is_nan(v) else v
    # επίσης βγάλε και “friendly” percent keys (π.χ. voice_vs_target -> voice_vs_target_pct)
    for k, v in row_dict.items():
        if k.endswith("_vs_target"):
            out[f"{k}_pct"] = format_percent(v)
    return out

# ───────────────────────────── UI ─────────────────────────────
st.title(APP_TITLE)

left, right = st.columns([2, 1])

with left:
    st.subheader("1) Templates & Mapping")
    tpl_bex_up = st.file_uploader("BEX template (.docx)", type=["docx"], key="tpl_bex")
    tpl_nonbex_up = st.file_uploader("NON-BEX template (.docx)", type=["docx"], key="tpl_nonbex")
    default_up = st.file_uploader("Default template για όλους (.docx) — προαιρετικό", type=["docx"], key="tpl_default")

    tpl_bex_path = tpl_nonbex_path = uploaded_default_path = None
    if tpl_bex_up:
        (RUNTIME / "bex.docx").write_bytes(tpl_bex_up.getvalue())
        tpl_bex_path = RUNTIME / "bex.docx"
        st.success("✔ Φορτώθηκε BEX template")
    if tpl_nonbex_up:
        (RUNTIME / "nonbex.docx").write_bytes(tpl_nonbex_up.getvalue())
        tpl_nonbex_path = RUNTIME / "nonbex.docx"
        st.success("✔ Φορτώθηκε NON-BEX template")
    if default_up:
        (RUNTIME / "default_uploaded.docx").write_bytes(default_up.getvalue())
        uploaded_default_path = RUNTIME / "default_uploaded.docx"
        st.success("✔ Φορτώθηκε Default template")

    map_up = st.file_uploader("store_mapping.json (προαιρετικό — ονόματα/κατηγορία/template ανά store)", type=["json"])
    if map_up:
        (RUNTIME / "store_mapping.json").write_bytes(map_up.getvalue())
        mapping_path = RUNTIME / "store_mapping.json"
        st.info("Χρησιμοποιείται το ανεβασμένο store_mapping.json (runtime).")
    elif REPO_MAPPING.exists():
        mapping_path = REPO_MAPPING
        st.info("Χρησιμοποιείται store_mapping.json από το repo.")
    else:
        mapping_path = None
        st.caption("Δεν υπάρχει store_mapping.json — όλα τα stores θα πάνε στο NON-BEX, εκτός αν ορίσεις λίστα BEX παρακάτω.")

with right:
    st.subheader("BEX detection")
    bex_mode = st.radio("Πώς βρίσκουμε αν είναι BEX;", ["Από λίστα", "Από στήλη (YES/NO)"], index=0)
    bex_list = set()
    bex_col = ""
    if bex_mode == "Από λίστα":
        bex_input = st.text_area("BEX stores (comma-separated)", "DRZ01, FKM01, ESC01, LND01, PKK01")
        bex_list = {s.strip().upper() for s in bex_input.split(",") if s.strip()}
    else:
        bex_col = st.text_input("Όνομα στήλης στο Excel που έχει YES/NO για BEX", "bex_store")

# ── Excel upload ──
st.subheader("2) Ανέβασε Excel")
if pd is None:
    st.error("Χρειάζονται pandas και openpyxl στο requirements.txt")
    st.stop()

xls = st.file_uploader("Excel (.xlsx)", type=["xlsx"])
sheet_name = st.text_input("Sheet (προαιρετικό — κενό για 1ο sheet)", value="")
df = None

if xls is not None:
    try:
        df = pd.read_excel(xls, sheet_name=sheet_name or 0)
        orig_cols = list(df.columns)
        norm_cols = [_norm_header(c) for c in df.columns]
        df.columns = norm_cols

        # alias για store_code
        aliases = ["store", "storeid", "store_id", "code", "dealer", "dealerid", "dealer_id", "dealercode", "dealer_code"]
        if "store_code" not in df.columns:
            for a in aliases:
                if a in df.columns:
                    df.rename(columns={a: "store_code"}, inplace=True)
                    break

        st.markdown("**Headers (original):**")
        st.code(str(orig_cols))
        st.markdown("**Headers (normalized):**")
        st.code(str(list(df.columns)))

        st.success(f"OK: {len(df)} γραμμές, {len(df.columns)} στήλες.")
        st.dataframe(df.head(15), use_container_width=True)

    except Exception as e:
        st.error(f"Σφάλμα ανάγνωσης Excel: {e}")

# ── Template audit (προαιρετικό αλλά χρήσιμο) ──
st.subheader("Template audit (placeholders που βρέθηκαν στα .docx)")
audit_cols = st.columns(3)
with audit_cols[0]:
    if tpl_bex_path:
        doc = Document(str(tpl_bex_path))
        st.caption("BEX template placeholders:")
        st.code(sorted(extract_placeholders_from_doc(doc)))
with audit_cols[1]:
    if tpl_nonbex_path:
        doc = Document(str(tpl_nonbex_path))
        st.caption("NON-BEX template placeholders:")
        st.code(sorted(extract_placeholders_from_doc(doc)))
with audit_cols[2]:
    if uploaded_default_path:
        doc = Document(str(uploaded_default_path))
        st.caption("Default template placeholders:")
        st.code(sorted(extract_placeholders_from_doc(doc)))

# ── Generate ──
st.subheader("3) Παραγωγή ανά κατάστημα & λήψη ZIP")
go = st.button("🚀 Generate")

if go:
    if df is None or df.empty:
        st.error("Λείπουν δεδομένα Excel.")
        st.stop()

    # φόρτωσε mapping (αν υπάρχει)
    store_map = load_store_mapping(mapping_path)

    # In-memory zip
    out_buf = io.BytesIO()
    zf = zipfile.ZipFile(out_buf, "w", compression=zipfile.ZIP_DEFLATED)
    generated: list[str] = []
    errors: list[Tuple[str, str]] = []

    for _, row in df.iterrows():
        row_dict = {k: ("" if _is_nan(v) else v) for k, v in row.to_dict().items()}
        store_code = _safe_str(row_dict.get("store_code")).upper().strip()
        if not store_code:
            errors.append(("[missing store_code]", "Άδεια τιμή store_code"))
            continue

        # mapping.json info (προαιρετικά)
        info = store_map.get(store_code, store_map.get("_default", {}))
        store_name = info.get("store_name", store_code)
        category = info.get("category", "NON_BEX")
        store_template_name = info.get("template", "default.docx")

        # override κατηγορίας από Excel αν έχεις στήλη category
        if "category" in row_dict and str(row_dict["category"]).strip():
            category = str(row_dict["category"]).strip()

        # ή από BEX detection σύμφωνα με UI
        if bex_mode == "Από λίστα":
            if store_code in bex_list:
                category = "BEX"
            else:
                category = "NON_BEX"
        else:  # Από στήλη YES/NO
            flag = str(row_dict.get(_norm_header(bex_col), "")).strip().lower()
            category = "BEX" if flag in {"yes", "y", "1", "true", "ναι"} else "NON_BEX"

        # επίλεξε template με ιεραρχία (uploaded default -> category -> per-store -> repo default)
        chosen_tpl = pick_template_path(
            store_template_name,
            category,
            uploaded_default_path,
            tpl_bex_path,
            tpl_nonbex_path,
        )
        if not chosen_tpl.exists():
            errors.append((store_code, f"Template δεν βρέθηκε: {chosen_tpl}"))
            continue

        # χτίσε mapping για placeholders
        placeholders = build_placeholder_map(store_code, store_name, row_dict)

        try:
            doc = Document(str(chosen_tpl))
            replace_all(doc, placeholders)
            subdir = "BEX" if str(category).upper() == "BEX" else "NON_BEX"
            out_name = f"{subdir}/Letter_{store_code}.docx"

            mem = io.BytesIO()
            doc.save(mem)
            zf.writestr(out_name, mem.getvalue())
            generated.append(out_name)
        except Exception as e:
            errors.append((store_code, f"Docx error: {e}"))

    zf.close()
    out_buf.seek(0)

    if generated:
        st.success(f"Δημιουργήθηκαν {len(generated)} αρχεία.")
        st.download_button(
            "⬇️ Κατέβασε ZIP",
            data=out_buf.getvalue(),
            file_name="Nova_Letters_BEX_NONBEX.zip",
            mime="application/zip",
        )
        with st.expander("Περιεχόμενα ZIP"):
            st.write("\n".join(generated))
    if errors:
        st.error("Αποτυχίες:")
        for s, msg in errors:
            st.write("•", s, "→", msg)