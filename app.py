
"""
Nova Letters & Templates Generator (stable fallback build)
- Reads store_mapping.json
- Loads the correct Word template from ./templates/
- Replaces [[placeholders]] safely (paragraphs & tables)
- Accepts metrics from JSON or Excel (by column HEADERS, not letters)
- Formats percentages: 1.22 -> 122%, 0.85 -> 85%, 122 -> 122%
Usage (CLI):
  python app.py --store FKM01 --data data.json --out out_dir
  python app.py --store FKM01 --excel metrics.xlsx --sheet Sheet1 --out out_dir
"""

import os, sys, json, argparse, datetime
from typing import Any, Dict
try:
    import pandas as pd
except Exception:
    pd = None

from docx import Document

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATES_DIR = os.path.join(HERE, "templates")
MAPPING_FILE = os.path.join(HERE, "store_mapping.json")

# ---------- helpers ----------

def format_percent(x: Any) -> str:
    """Formats a value to a percentage string.
    Rules:
    - if 0 <= x < 1  -> x*100 (0.85 -> 85%)
    - if 1 <= x < 10 -> x*100 (1.22 -> 122%)
    - if x >= 10     -> assume already in percent (122 -> 122%)
    Non-numeric -> returns as-is.
    """
    try:
        val = float(x)
    except Exception:
        return str(x)
    if val < 1:
        return f"{val*100:.0f}%"
    if val < 10:
        return f"{val*100:.0f}%"
    return f"{val:.0f}%"
    
def replace_all(doc: Document, mapping: Dict[str, str]) -> None:
    """Replace [[placeholder]] in paragraphs and tables."""
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

def load_store_mapping() -> Dict[str, Any]:
    if not os.path.exists(MAPPING_FILE):
        raise FileNotFoundError(f"Missing mapping file: {MAPPING_FILE}")
    with open(MAPPING_FILE, "r", encoding="utf-8") as f:
        return json.load(f)

def select_template(store_code: str, store_map: Dict[str, Any]) -> (str, str):
    info = store_map.get(store_code, store_map.get("_default", {}))
    template_name = info.get("template", "default.docx")
    store_name = info.get("store_name", store_code)
    template_path = os.path.join(TEMPLATES_DIR, template_name)
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Template not found: {template_path}")
    return template_path, store_name

def read_metrics_from_excel(path: str, sheet: str) -> Dict[str, Any]:
    if pd is None:
        raise RuntimeError("pandas is required for Excel. Install pandas/openpyxl.")
    df = pd.read_excel(path, sheet_name=sheet)
    # Expect headers matching keys; we take the first row.
    row = df.iloc[0].to_dict()
    return {k: row.get(k) for k in df.columns}

def build_placeholder_map(store_code: str, store_name: str, payload: Dict[str, Any]) -> Dict[str, Any]:
    today = datetime.date.today()
    mm = today.strftime("%B")
    placeholders = {
        "store_code": store_code,
        "store_name": store_name,
        "month_name": mm,
        "year": today.year,
        "comment": payload.get("comment", ""),
        # numbers:
        "fixed_target": payload.get("fixed_target", ""),
        "fixed_actual": payload.get("fixed_actual", ""),
        "ftth_actual": payload.get("ftth_actual", ""),
        "eon_tv_actual": payload.get("eon_tv_actual", ""),
        "mobile_upgrades": payload.get("mobile_upgrades", ""),
        "pending_mobile": payload.get("pending_mobile", ""),
    }
    # Percent-derived:
    placeholders["voice_vs_target_pct"] = format_percent(payload.get("voice_vs_target", ""))
    return placeholders

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--store", required=True, help="Store code, e.g., FKM01")
    ap.add_argument("--data", help="Path to JSON with metrics")
    ap.add_argument("--excel", help="Path to Excel with headers matching keys")
    ap.add_argument("--sheet", default=None, help="Excel sheet name")
    ap.add_argument("--out", default="out", help="Output directory")
    args = ap.parse_args()

    store_map = load_store_mapping()
    template_path, store_name = select_template(args.store, store_map)

    # load payload
    payload = {}
    if args.data and os.path.exists(args.data):
        with open(args.data, "r", encoding="utf-8") as f:
            payload = json.load(f)
    elif args.excel and os.path.exists(args.excel):
        payload = read_metrics_from_excel(args.excel, args.sheet or 0)
    else:
        print("No data source provided; using empty placeholders.")

    placeholders = build_placeholder_map(args.store, store_name, payload)

    # generate
    os.makedirs(args.out, exist_ok=True)
    out_path = os.path.join(args.out, f"Letter_{args.store}.docx")
    doc = Document(template_path)
    replace_all(doc, placeholders)
    doc.save(out_path)
    print(f"OK: created {out_path}")

if __name__ == "__main__":
    main()
