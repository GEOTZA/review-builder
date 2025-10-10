
import streamlit as st
import re, os, io, json, zipfile
from typing import Dict, Any, List, Tuple
from docx import Document
from docx.oxml.ns import qn

# PDF text extraction
try:
    from pdfminer.high_level import extract_text
except Exception as e:
    extract_text = None

st.set_page_config(page_title="Review/Plan Generator", layout="wide")

def set_default_font(doc, font_name="Aptos"):
    for style in doc.styles:
        if hasattr(style, "font"):
            try:
                style.font.name = font_name
                style._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
                style._element.rPr.rFonts.set(qn('w:cs'), font_name)
            except Exception:
                pass

def read_pdf_text(file_bytes: bytes) -> str:
    if extract_text is None:
        raise RuntimeError("pdfminer.six is not installed on this environment.")
    # pdfminer expects a file-like path or descriptor; wrap bytes in BytesIO and use extract_text
    bio = io.BytesIO(file_bytes)
    text = extract_text(bio) or ""
    return text

def parse_metrics(text: str, patterns: Dict[str, str]) -> Dict[str, Any]:
    results: Dict[str, Any] = {}
    for key, pat in patterns.items():
        try:
            m = re.search(pat, text, re.IGNORECASE | re.MULTILINE)
        except re.error as e:
            results[key] = f"[Regex error: {e}]"
            continue
        if m:
            # prefer the last capturing group if there are multiple
            group_val = None
            if m.lastindex and m.lastindex >= 1:
                group_val = m.group(m.lastindex)
            else:
                group_val = m.group(1) if len(m.groups()) >= 1 else m.group(0)
            val = (group_val or "").strip()
            val = val.replace(",", ".")
            # Try float cast but keep original if not numeric
            try:
                num = float(re.sub(r"[^0-9.\-]", "", val))
                results[key] = num
            except ValueError:
                results[key] = val
        else:
            results[key] = ""
    return results

def replace_placeholders(doc: Document, mapping: Dict[str, Any]):
    def repl_text(s: str) -> str:
        def rfun(m):
            k = m.group(1)
            return str(mapping.get(k, ""))
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

st.title("📄 Review/Plan Generator (BEX & Non‑BEX)")

with st.sidebar:
    st.header("⚙️ Settings")
    st.caption("Οδήγησε το extraction με BEX λίστα και regex patterns.")
    bex_input = st.text_area("BEX stores (comma-separated)", "ESC01,FKM01,LND01,DRZ01,PKK01")
    bex_set = set([s.strip().upper() for s in bex_input.split(",") if s.strip()])

    default_review_patterns = {
        "mobile_actual": r"Κινητ(ή|ης).*?(\d+)\s*(ενεργοποιήσεις|γραμμές)",
        "fixed_actual":  r"Σταθερ(ή|ής).*?(\d+)\s*(ενεργοποιήσεις|γραμμές)",
        "ftth_actual":   r"FTTH.*?(\d+)",
        "fwa_actual":    r"FWA.*?(\d+)",
        "eon_actual":    r"(EON|TV).*?(\d+)"
    }
    default_plan_patterns = {
        "mobile_target": r"(?i)κινητ[ήής].*?(\d+)\s*(?:γρ|lines|γραφ|ενεργ)",
        "fixed_target":  r"(?i)σταθερ[ήής].*?(\d+)\s*(?:γρ|lines|γραφ|ενεργ)",
        "ftth_target":   r"(?i)ftth.*?(\d+)",
        "fwa_target":    r"(?i)fwa.*?(\d+)",
        "eon_target":    r"(?i)(eon|tv).*?(\d+)"
    }
    review_patterns_json = st.text_area("Review regex (JSON)", json.dumps(default_review_patterns, ensure_ascii=False, indent=2))
    plan_patterns_json   = st.text_area("Plan regex (JSON)", json.dumps(default_plan_patterns, ensure_ascii=False, indent=2))
    try:
        review_patterns = json.loads(review_patterns_json)
        plan_patterns = json.loads(plan_patterns_json)
    except Exception as e:
        st.error(f"JSON error: {e}")
        st.stop()

    st.markdown("---")
    st.subheader("📄 Templates (.docx)")
    tpl_bex = st.file_uploader("BEX template", type=["docx"], key="tpl_bex")
    tpl_nonbex = st.file_uploader("Non‑BEX template", type=["docx"], key="tpl_nonbex")
    st.caption("Χρησιμοποίησε placeholders όπως [[review_mobile_actual]], [[plan_mobile_target]], [[title]], [[store]].")

st.markdown("### 1) Ανέβασε τα PDFs σου")
st.caption("Ονόμασε τα αρχεία ως <STORE>_Review.pdf και <STORE>_Plan.pdf (π.χ. ESC01_Review.pdf, ESC01_Plan.pdf).")
uploaded = st.file_uploader("PDFs", type=["pdf"], accept_multiple_files=True)

st.markdown("### 2) Δώσε λίστα καταστημάτων")
store_list_text = st.text_area("STORE codes (ένα ανά γραμμή)", "ESC01\nPKK01")

run = st.button("🔧 Parse & Generate")

if run:
    if not uploaded:
        st.error("Ανέβασε πρώτα τα PDFs.")
        st.stop()
    if not tpl_bex or not tpl_nonbex:
        st.error("Ανέβασε και τα δύο templates (BEX & Non‑BEX).")
        st.stop()
    if extract_text is None:
        st.error("Λείπει η βιβλιοθήκη pdfminer.six. Τρέξε: pip install pdfminer.six")
        st.stop()

    # Index uploaded files by name
    pdf_map = {f.name: f for f in uploaded}
    stores = [s.strip().upper() for s in store_list_text.splitlines() if s.strip()]

    # Load templates in memory
    tpl_bex_bytes = tpl_bex.read()
    tpl_nonbex_bytes = tpl_nonbex.read()

    # Collect results to zip
    out_zip = io.BytesIO()
    z = zipfile.ZipFile(out_zip, mode="w", compression=zipfile.ZIP_DEFLATED)

    summary_rows: List[Tuple[str, Dict[str, Any], Dict[str, Any]]] = []

    for code in stores:
        r_name = f"{code}_Review.pdf"
        p_name = f"{code}_Plan.pdf"

        if r_name not in pdf_map or p_name not in pdf_map:
            st.warning(f"{code}: λείπει κάποιο PDF ({r_name} ή {p_name})")
            continue

        # Extract text
        review_text = read_pdf_text(pdf_map[r_name].read())
        plan_text   = read_pdf_text(pdf_map[p_name].read())

        # Parse metrics
        review_vals = parse_metrics(review_text, review_patterns)
        plan_vals   = parse_metrics(plan_text, plan_patterns)

        # Merge mapping
        mapping = {}
        mapping.update({f"review_{k}": v for k, v in review_vals.items()})
        mapping.update({f"plan_{k}": v for k, v in plan_vals.items()})
        mapping["store"] = code
        mapping["title"] = f"Review September 2025 — Plan October 2025 — {code}"

        # Choose template
        is_bex = (code in bex_set)
        tpl_bytes = tpl_bex_bytes if is_bex else tpl_nonbex_bytes
        # Build Word
        doc = Document(io.BytesIO(tpl_bytes))
        set_default_font(doc, "Aptos")
        replace_placeholders(doc, mapping)

        # Save to zip
        out_name = f"{code}_ReviewSep_PlanOct.docx"
        out_buf = io.BytesIO()
        doc.save(out_buf)
        z.writestr(out_name, out_buf.getvalue())

        summary_rows.append((code, review_vals, plan_vals))

    z.close()
    st.success("Ολοκληρώθηκε η δημιουργία των αρχείων.")
    st.download_button("⬇️ Κατέβασε όλα τα .docx (ZIP)", data=out_zip.getvalue(), file_name="reviews_plan_docs.zip")

    # Show a summary table
    if summary_rows:
        st.markdown("### Σύνοψη εξαγόμενων τιμών")
        for code, r, p in summary_rows:
            with st.expander(code, expanded=False):
                st.write("**Review values**", r)
                st.write("**Plan values**", p)
