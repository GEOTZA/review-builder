import streamlit as st
import re, os, io, json, zipfile
from typing import Dict, Any, List, Tuple
from docx import Document
from docx.oxml.ns import qn

# PDF text extraction
try:
    from pdfminer.high_level import extract_text
except Exception:
    extract_text = None

st.set_page_config(page_title="Review/Plan Generator", layout="wide")

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

def read_pdf_text_from_bytes(file_bytes: bytes) -> str:
    """Return extracted text or empty string; never raises to the UI."""
    if extract_text is None:
        return ""
    try:
        bio = io.BytesIO(file_bytes)
        txt = extract_text(bio) or ""
        return txt
    except Exception:
        return ""

def parse_metrics(text: str, patterns: Dict[str, str]) -> Dict[str, Any]:
    """
    Extract values using regex patterns; robust to errors.
    Keeps values as strings (no forced float cast).
    """
    results: Dict[str, Any] = {}
    for key, pat in patterns.items():
        try:
            m = re.search(pat, text, re.IGNORECASE | re.MULTILINE | re.DOTALL)
        except re.error as e:
            results[key] = f"[Regex error: {e}]"
            continue
        if m:
            try:
                # prefer last capturing group if multiple
                val = (m.group(m.lastindex) if m.lastindex else m.group(1))
            except Exception:
                val = m.group(0)
            val = (val or "").strip()
            val = val.replace(",", ".")
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

def parse_file_role(name: str):
    """
    Χαλαρή αναγνώριση STORE + ρόλου (review/plan) από filename.
    Δέχεται:
      <STORE>_Business Review.pdf
      <STORE>_Action Plan.pdf
      <STORE>-Business-Review.pdf
      <STORE> BusinessReview.pdf
      <STORE>_review.pdf
      <STORE>_plan.pdf
      <STORE>_ActionPlan.pdf
      <STORE>_BusinessReview.pdf
    """
    base = name.strip()
    if base.lower().endswith(".pdf"):
        base = base[:-4]
    # συμπύκνωσε πολλαπλά κενά
    base = re.sub(r"\s+", " ", base)

    # Πάρε το STORE ως πρώτο κομμάτι
    m_store = re.match(r"^\s*([A-Za-z0-9]+)[\s_\-]+(.+)$", base)
    if not m_store:
        only_code = re.match(r"^\s*([A-Za-z0-9]+)\s*$", base)
        if only_code:
            return only_code.group(1).upper(), None
        return None, None
    store = m_store.group(1).upper()
    tail = m_store.group(2).strip().lower()

    # compact για εύκολα matches
    compact = re.sub(r"[\s_\-]+", "", tail)

    # review keys
    review_keys = ["businessreview", "review"]
    # plan keys
    plan_keys = ["actionplan", "plan"]

    if any(k in compact for k in review_keys) or tail.endswith("review"):
        return store, "review"
    if any(k in compact for k in plan_keys) or tail.endswith("plan"):
        return store, "plan"

    # fallback
    if "review" in compact:
        return store, "review"
    if "plan" in compact:
        return store, "plan"
    return store, None

# ---------- UI ----------
st.title("📄 Review/Plan Generator (BEX & Non-BEX)")

with st.sidebar:
    st.header("⚙️ Settings")
    st.caption("Ορίσε BEX λίστα & regex extraction (JSON).")
    bex_input = st.text_area("BEX stores (comma-separated)", "ESC01,FKM01,LND01,DRZ01,PKK01")
    bex_set = set([s.strip().upper() for s in bex_input.split(",") if s.strip()])

    default_review_patterns = {
        "mobile_actual": r"Κινητ(ή|ης).*?(\d+)\s*(ενεργοποιήσεις|γραμμές)",
        "fixed_actual":  r"Σταθερ(ή|ής).*?(\d+)\s*(ενεργοποιήσεις|γραμμές)",
        "ftth_actual":   r"FTTH.*?(\d+)",
        "fwa_actual":    r"FWA.*?(\d+)",
        "eon_actual":    r"(EON|TV).*?(\d+)",
        # μπορείς να προσθέσεις κι άλλα, π.χ. pending:
        # "pending_mobile": r"pending.*?κινητ.*?(\d+)"
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

    st.subheader("📄 Templates (.docx)")
    tpl_bex = st.file_uploader("BEX template", type=["docx"], key="tpl_bex")
    tpl_nonbex = st.file_uploader("Non-BEX template", type=["docx"], key="tpl_nonbex")
    st.caption("Placeholders: [[review_mobile_actual]], [[plan_mobile_target]], [[title]], [[store]], κ.λπ.")

st.markdown("### 1) Ανέβασε PDFs")
st.caption("Δεκτές ονομασίες (χαλαρό matching): <STORE>_Business Review.pdf & <STORE>_Action Plan.pdf, αλλά και ESC01_Review.pdf/ESC01_Plan.pdf, με ή χωρίς παύλες/κενά.")
uploaded = st.file_uploader("PDFs", type=["pdf"], accept_multiple_files=True)

st.markdown("### 2) Δώσε λίστα καταστημάτων")
store_list_text = st.text_area("STORE codes (ένα ανά γραμμή)", "ESC01\nLCI01\nLGS01\nLND01")

run = st.button("🔧 Parse & Generate")

if run:
    if not uploaded:
        st.error("Ανέβασε πρώτα τα PDFs.")
        st.stop()
    if not tpl_bex or not tpl_nonbex:
        st.error("Ανέβασε και τα δύο templates (.docx).")
        st.stop()
    if extract_text is None:
        st.error("Λείπει pdfminer.six στο περιβάλλον.")
        st.stop()

    # index uploaded files
    pdf_map = {}
    unmatched = []
    for f in uploaded:
        store, role = parse_file_role(f.name)
        if not store or not role:
            unmatched.append(f.name)
            continue
        pdf_map.setdefault(store, {})[role] = f
    if unmatched:
        st.warning("Μη αναγνωρισμένα filenames: " + ", ".join(unmatched))

    # load templates
    tpl_bex_bytes = tpl_bex.read()
    tpl_nonbex_bytes = tpl_nonbex.read()

    # prepare zip
    out_zip = io.BytesIO()
    z = zipfile.ZipFile(out_zip, "w", zipfile.ZIP_DEFLATED)

    stores = [s.strip().upper() for s in store_list_text.splitlines() if s.strip()]
    built = 0

    for code in stores:
        cols = st.columns([1,2,2])
        cols[0].markdown(f"**{code}**")

        pair = pdf_map.get(code, {})
        rfile = pair.get("review")
        pfile = pair.get("plan")
        if not rfile or not pfile:
            cols[1].warning("Λείπει Business Review ή Action Plan PDF.")
            continue

        rbytes = rfile.read()
        pbytes = pfile.read()

        rtext = read_pdf_text_from_bytes(rbytes)
        ptext = read_pdf_text_from_bytes(pbytes)

        cols[1].write(f"Review text chars: {len(rtext)}")
        cols[2].write(f"Plan text chars: {len(ptext)}")

        if len(rtext) < 20 or len(ptext) < 20:
            cols[1].error("Ένα από τα PDF φαίνεται σαν εικόνα (δεν έχει κείμενο). Κάνε OCR (Recognize Text) και ξαναδοκίμασε.")
            continue

        rvals = parse_metrics(rtext, review_patterns)
        pvals = parse_metrics(ptext, plan_patterns)

        with st.expander(f"Diagnostics — {code}", expanded=False):
            st.write("Review extracted:", rvals)
            st.write("Plan extracted:", pvals)

        mapping = {}
        mapping.update({f"review_{k}": v for k, v in rvals.items()})
        mapping.update({f"plan_{k}": v for k, v in pvals.items()})
        mapping["store"] = code
        mapping["title"] = f"Review September 2025 — Plan October 2025 — {code}"

        # choose template
        is_bex = (code in bex_set)
        tpl_bytes = tpl_bex_bytes if is_bex else tpl_nonbex_bytes
        doc = Document(io.BytesIO(tpl_bytes))
        set_default_font(doc, "Aptos")
        replace_placeholders(doc, mapping)

        out_name = f"{code}_ReviewSep_PlanOct.docx"
        buf = io.BytesIO()
        doc.save(buf)
        z.writestr(out_name, buf.getvalue())
        built += 1

    z.close()
    if built == 0:
        st.error("Δεν δημιουργήθηκε αρχείο. Έλεγξε filenames, ότι υπάρχουν και τα 2 PDFs ανά store και ότι δεν είναι σκαναρισμένες εικόνες.")
    else:
        st.success(f"Ολοκληρώθηκε: δημιουργήθηκαν {built} αρχεία.")
        st.download_button("⬇️ Κατέβασε όλα τα .docx (ZIP)", data=out_zip.getvalue(), file_name="reviews_plan_docs.zip")
