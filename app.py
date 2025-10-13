import streamlit as st
from pathlib import Path
import json, datetime, io, zipfile
from typing import Any, Dict

try:
    import pandas as pd
except Exception:
    pd = None

from docx import Document

st.set_page_config(page_title="Nova Letters — Batch Builder", layout="wide")

HERE = Path(__file__).parent
RUNTIME = HERE / "runtime"
RUNTIME.mkdir(exist_ok=True)
TEMPLATES_DIR = HERE / "templates"
DEFAULT_TEMPLATE = TEMPLATES_DIR / "default.docx"
REPO_MAPPING = HERE / "store_mapping.json"

# ---------- Helpers ----------
def format_percent(x: Any) -> str:
    try:
        val = float(x)
    except Exception:
        return str(x)
    if val < 1:
        return f"{val*100:.0f}%"
    if val < 10:
        return f"{val*100:.0f}%"
    return f"{val:.0f}%"

def replace_all(doc: Document, mapping: Dict[str, Any]) -> None:
    def repl_text(text: str) -> str:
        out = text
        for k, v in mapping.items():
            out = out.replace(f"[[{k}]]", str(v))
        return out

    for p in doc.paragraphs:
        p.text = repl_text(p.text)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                cell.text = repl_text(cell.text)

def load_store_mapping(path: Path) -> Dict[str, Any]:
    with path.open("r", encoding="utf-8") as f:
        return json.load(f)

def build_placeholder_map(store_code: str, store_name: str, payload: Dict[str, Any]) -> Dict[str, Any]:
    today = datetime.date.today()
    out = {
        "store_code": store_code,
        "store_name": store_name,
        "month_name": today.strftime("%B"),
        "year": today.year,
        "comment": payload.get("comment", ""),
        "fixed_target": payload.get("fixed_target", ""),
        "fixed_actual": payload.get("fixed_actual", ""),
        "ftth_actual": payload.get("ftth_actual", ""),
        "eon_tv_actual": payload.get("eon_tv_actual", ""),
        "mobile_upgrades": payload.get("mobile_upgrades", ""),
        "pending_mobile": payload.get("pending_mobile", ""),
        "voice_vs_target_pct": format_percent(payload.get("voice_vs_target", "")),
    }
    # Επίσης περνάμε *όλα* τα υπόλοιπα columns ως [[column_name]]
    for k, v in payload.items():
        if k not in out:
            out[k] = v
    return out

def pick_template_path(template_name: str, uploaded_template: Path | None) -> Path:
    # Αν υπάρχει uploaded custom template, δώσε προτεραιότητα
    if uploaded_template and uploaded_template.exists():
        return uploaded_template
    # αλλιώς ψάξε στο templates/
    cand = TEMPLATES_DIR / (template_name or "default.docx")
    return cand if cand.exists() else DEFAULT_TEMPLATE

# ---------- UI ----------
st.title("📄 Nova Letters — Μαζική Παραγωγή (BEX / NON-BEX)")

# Mapping & Template
st.subheader("1) Mapping & Template")

c1, c2 = st.columns(2)

with c1:
    st.markdown("**store_mapping.json** (repo ή ανέβασέ το)")
    m_up = st.file_uploader("Upload store_mapping.json", type=["json"])
    if m_up:
        (RUNTIME / "store_mapping.json").write_bytes(m_up.getvalue())
        st.success("Uploaded to runtime/store_mapping.json")
    if (RUNTIME / "store_mapping.json").exists():
        mapping_path = RUNTIME / "store_mapping.json"
        st.info("Using uploaded mapping (runtime).")
    elif REPO_MAPPING.exists():
        mapping_path = REPO_MAPPING
        st.info("Using mapping from repo.")
    else:
        mapping_path = None
        st.error("Missing store_mapping.json")

with c2:
    st.markdown("**Template (.docx)** (repo ή ανέβασέ το)")
    t_up = st.file_uploader("Upload template .docx", type=["docx"])
    uploaded_template = None
    if t_up:
        (RUNTIME / "custom_template.docx").write_bytes(t_up.getvalue())
        uploaded_template = RUNTIME / "custom_template.docx"
        st.success("Uploaded to runtime/custom_template.docx")
    else:
        uploaded_template = None
    if DEFAULT_TEMPLATE.exists():
        st.info("Repo default template: templates/default.docx")

st.subheader("2) Δεδομένα — Excel (1 γραμμή ανά κατάστημα)")
if pd is None:
    st.error("Λείπει pandas/openpyxl. Πρόσθεσε στο requirements.txt: pandas, openpyxl")
    st.stop()

excel = st.file_uploader("Upload Excel", type=["xlsx", "xls"])
sheet = st.text_input("Sheet (optional)", value="")
df = None
if excel is not None:
    try:
        df = pd.read_excel(excel, sheet_name=sheet or 0)
        st.success(f"Φορτώθηκαν {len(df)} γραμμές από Excel.")
        st.write("**Preview των values που θα περάσουν:**")
        st.dataframe(df, use_container_width=True)
    except Exception as e:
        st.error(f"Σφάλμα ανάγνωσης Excel: {e}")

st.subheader("3) Παραγωγή ανά κατάστημα & ομαδοποίηση")
start = st.button("🚀 Generate BEX / NON-BEX & κατέβασέ τα σε ZIP")

if start:
    if mapping_path is None:
        st.error("Λείπει store_mapping.json")
        st.stop()
    if df is None or df.empty:
        st.error("Λείπουν δεδομένα Excel.")
        st.stop()

    try:
        store_map = load_store_mapping(mapping_path)
    except Exception as e:
        st.error(f"Δεν διαβάζεται το mapping: {e}")
        st.stop()

    # In-memory zip
    mem_zip = io.BytesIO()
    zf = zipfile.ZipFile(mem_zip, mode="w", compression=zipfile.ZIP_DEFLATED)

    generated = []
    errors = []

    for idx, row in df.iterrows():
        row_dict = {k: ("" if pd.isna(v) else v) for k, v in row.to_dict().items()}

        store_code = str(row_dict.get("store_code", "")).strip()
        if not store_code:
            errors.append((idx, "Missing store_code"))
            continue

        info = store_map.get(store_code, store_map.get("_default", {}))
        store_name = info.get("store_name", store_code)
        template_name = info.get("template", "default.docx")
        category = info.get("category", "NON_BEX")  # default grouping

        # template pick
        t_path = pick_template_path(template_name, uploaded_template)
        if not t_path or not t_path.exists():
            errors.append((store_code, f"Template not found: {t_path}"))
            continue

        # placeholders
        placeholders = build_placeholder_map(store_code, store_name, row_dict)

        # build doc
        try:
            doc = Document(str(t_path))
            replace_all(doc, placeholders)
            # save to zip path: BEX/Letter_FKM01.docx ή NON_BEX/Letter_DRZ01.docx
            subdir = "BEX" if str(category).upper() == "BEX" else "NON_BEX"
            out_name = f"{subdir}/Letter_{store_code}.docx"
            buf = io.BytesIO()
            doc.save(buf)
            buf.seek(0)
            zf.writestr(out_name, buf.read())
            generated.append(out_name)
        except Exception as e:
            errors.append((store_code, str(e)))

    zf.close()
    mem_zip.seek(0)

    if generated:
        st.success(f"Δημιουργήθηκαν {len(generated)} αρχεία.")
        st.download_button(
            "⬇️ Κατέβασε ZIP (BEX/NON-BEX)",
            data=mem_zip,
            file_name="Nova_Letters_BEX_NONBEX.zip",
            mime="application/zip",
        )
        with st.expander("Δείτε τα αρχεία που περιέχονται"):
            st.write("\n".join(generated))
    if errors:
        st.error("Κάποια stores δεν δημιουργήθηκαν:")
        for e in errors:
            st.write("•", e)