# app.py — Streamlit έκδοση (runs in Streamlit Cloud)
import streamlit as st
from pathlib import Path
import json, datetime, io
from typing import Any, Dict

try:
    import pandas as pd
except Exception:
    pd = None

from docx import Document

st.set_page_config(page_title="Nova Letters Generator", layout="wide")

HERE = Path(__file__).parent
RUNTIME = HERE / "runtime"
RUNTIME.mkdir(exist_ok=True)
TEMPLATES_DIR = HERE / "templates"
DEFAULT_TEMPLATE = TEMPLATES_DIR / "default.docx"
REPO_MAPPING = HERE / "store_mapping.json"
REPO_DATA = HERE / "data.json"

# ---------------- helpers ----------------
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

def build_placeholder_map(store_code: str, store_name: str, payload: Dict[str, Any]) -> Dict[str, Any]:
    today = datetime.date.today()
    return {
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

def load_store_mapping(path: Path) -> Dict[str, Any]:
    with path.open("r", encoding="utf-8") as f:
        return json.load(f)

def read_metrics_from_excel(buf, sheet_name=None) -> Dict[str, Any]:
    if pd is None:
        raise RuntimeError("pandas/openpyxl not installed. Add them in requirements.txt")
    df = pd.read_excel(buf, sheet_name=sheet_name or 0)
    row = df.iloc[0].to_dict()
    return {k: row.get(k) for k in df.columns}

# ---------------- UI ----------------
st.title("📄 Nova Letters Generator")

# 1) Mapping & Template
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
    if t_up:
        (RUNTIME / "custom_template.docx").write_bytes(t_up.getvalue())
        st.success("Uploaded to runtime/custom_template.docx")
    if (RUNTIME / "custom_template.docx").exists():
        template_path = RUNTIME / "custom_template.docx"
        st.info("Using uploaded template (runtime).")
    elif DEFAULT_TEMPLATE.exists():
        template_path = DEFAULT_TEMPLATE
        st.info("Using templates/default.docx from repo.")
    else:
        template_path = None
        st.error("Missing template .docx")

# 2) Data
st.subheader("2) Δεδομένα")
tab_json, tab_excel, tab_form = st.tabs(["JSON", "Excel", "Φόρμα"])
payload = {}

with tab_json:
    d_up = st.file_uploader("Upload data.json (ή θα διαβαστεί του repo)", type=["json"])
    if d_up:
        payload = json.loads(d_up.getvalue().decode("utf-8"))
        st.success("Loaded data from uploaded JSON.")
    elif REPO_DATA.exists():
        payload = json.loads(REPO_DATA.read_text(encoding="utf-8"))
        st.info("Loaded data from repo/data.json.")
    else:
        st.warning("No JSON provided.")

with tab_excel:
    e_up = st.file_uploader("Upload Excel (headers = keys)", type=["xlsx", "xls"])
    sheet = st.text_input("Sheet name (optional)", value="")
    if e_up is not None:
        try:
            payload = read_metrics_from_excel(e_up, sheet or None)
            st.success("Loaded data from Excel (first row).")
        except Exception as e:
            st.error(f"Excel error: {e}")

with tab_form:
    with st.form("manual"):
        fixed_target = st.text_input("fixed_target")
        fixed_actual = st.text_input("fixed_actual")
        voice_vs_target = st.text_input("voice_vs_target (0.85 ή 1.22 ή 122)")
        ftth_actual = st.text_input("ftth_actual")
        eon_tv_actual = st.text_input("eon_tv_actual")
        mobile_upgrades = st.text_input("mobile_upgrades")
        pending_mobile = st.text_input("pending_mobile")
        comment = st.text_input("comment")
        ok = st.form_submit_button("Use form values")
    if ok:
        payload = {
            "fixed_target": fixed_target,
            "fixed_actual": fixed_actual,
            "voice_vs_target": voice_vs_target,
            "ftth_actual": ftth_actual,
            "eon_tv_actual": eon_tv_actual,
            "mobile_upgrades": mobile_upgrades,
            "pending_mobile": pending_mobile,
            "comment": comment,
        }
        st.success("Loaded data from form.")

# 3) Store & Generate
st.subheader("3) Κατάστημα & Παραγωγή")
store_code = st.text_input("Store code (π.χ. FKM01)", value="FKM01")

if st.button("Δημιούργησε Word"):
    if mapping_path is None:
        st.error("Λείπει store_mapping.json")
    elif template_path is None:
        st.error("Λείπει template .docx")
    elif not store_code.strip():
        st.error("Δώσε store code")
    else:
        try:
            store_map = load_store_mapping(mapping_path)
            info = store_map.get(store_code, store_map.get("_default", {}))
            template_name = info.get("template", "default.docx")
            store_name = info.get("store_name", store_code)

            # Αν mapping δείχνει άλλο template στον φάκελο templates/
            if template_path == DEFAULT_TEMPLATE and template_name != "default.docx":
                alt = TEMPLATES_DIR / template_name
                template_path_use = alt if alt.exists() else template_path
            else:
                template_path_use = template_path

            doc = Document(str(template_path_use))
            placeholders = build_placeholder_map(store_code, store_name, payload or {})
            replace_all(doc, placeholders)

            buf = io.BytesIO()
            doc.save(buf)
            buf.seek(0)

            fn = f"Letter_{store_code}.docx"
            st.success(f"Έτοιμο: {fn}")
            st.download_button("⬇️ Κατέβασε το Word", data=buf, file_name=fn,
                               mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        except Exception as e:
            st.error(f"Σφάλμα: {e}")
            st.stop()

st.caption("Placeholders στο template με μορφή [[placeholder]] π.χ. [[store_code]], [[voice_vs_target_pct]].")