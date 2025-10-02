#!/usr/bin/env python3
import argparse, os, io, re, unicodedata, traceback
import xml.etree.ElementTree as ET
from typing import Dict, List, Tuple, Optional, DefaultDict
from collections import defaultdict

# ========= Namespaces (Excel 2003 XML) =========
NAMESPACES = {
    "":     "urn:schemas-microsoft-com:office:spreadsheet",
    "o":    "urn:schemas-microsoft-com:office:office",
    "x":    "urn:schemas-microsoft-com:office:excel",
    "ss":   "urn:schemas-microsoft-com:office:spreadsheet",
    "html": "http://www.w3.org/TR/REC-html40",
}
for prefix, uri in NAMESPACES.items():
    ET.register_namespace(prefix, uri)

SS_NS = NAMESPACES["ss"]
NS = {"ss": SS_NS}

# ========= Unicode / whitespace normalization =========
SPACE_LIKES = [
    "\u00A0",  # NBSP
    "\u2000","\u2001","\u2002","\u2003","\u2004","\u2005",
    "\u2006","\u2007","\u2008","\u2009","\u200A",
    "\u202F",  # narrow NBSP
    "\u205F",  # medium math space
    "\u3000",  # ideographic space
]
ZERO_WIDTH = ["\u200B", "\u200C", "\u200D", "\uFEFF"]  # ZWSP/ZWNJ/ZWJ/BOM

def normalize_cell(s: str) -> str:
    if s is None:
        return ""
    s = unicodedata.normalize("NFKC", str(s))
    for ch in SPACE_LIKES:
        s = s.replace(ch, " ")
    for ch in ZERO_WIDTH:
        s = s.replace(ch, "")
    s = s.replace("\t", " ").replace("\r", "").replace("\n", " ")
    s = re.sub(r"\s+", " ", s)
    return s.strip()

# ========= Filename parsing (Section/panel and Gauge/color) =========
# Section: s + digits + ONE OR MORE letters (handles s8TIE, s1M, s5D, etc.)
SECTION_RE = re.compile(r"[sS]\d+[A-Za-z]+")

def parse_section_from_stem(stem: str) -> str:
    s = SECTION_RE.search(stem)
    return s.group(0) if s else "null"

import re

# Accept endings like:
#   ...14WHT
#   ...14WHT_01
#   ...14WHT_PLCIO
#   ...14WHT_PLCIO_01
GAUGE_COLOR_RE = re.compile(r"(?i)(\d{1,2})([A-Z]{3})(?:_PLCIO)?(?:_\d{2})?$")

def parse_gauge_color_from_stem(stem: str) -> str:
    """
    Extracts gauge+color (e.g., '14WHT') from the filename stem.
    Robust to optional suffixes (_PLCIO, _01, or both).
    Examples:
        's1MpA14WHT'             -> '14WHT'
        's5DpC1418WHT_PLCIO'     -> '18WHT'
        's5DpC1418WHT_PLCIO_02'  -> '18WHT'
        's3XpB16GRY_01'          -> '16GRY'
    """
    m = GAUGE_COLOR_RE.search(stem)
    if m:
        return f"{m.group(1)}{m.group(2).upper()}"
    # Fallback: find first occurrence anywhere (extra safety)
    m2 = re.search(r"(?i)(\d{1,2})([A-Z]{3})", stem)
    if m2:
        return f"{m2.group(1)}{m2.group(2).upper()}"
    return "unknown"

# ========= XML helpers =========

def parse_tree(path: str) -> ET.ElementTree:
    return ET.parse(path)

def find_table(root: ET.Element) -> Optional[ET.Element]:
    return root.find(".//ss:Table", NS)

def get_rows(table: ET.Element) -> List[ET.Element]:
    return table.findall("ss:Row", NS)

def enumerate_cells_with_positions(row: ET.Element):
    pos, out = 1, []
    for cell in row.findall("ss:Cell", NS):
        idx = cell.get(f"{{{SS_NS}}}Index")
        if idx:
            pos = int(idx)
        out.append((pos, cell))
        pos += 1
    return out

def get_cell_text(cell: ET.Element) -> str:
    d = cell.find("ss:Data", NS)
    return (d.text or "") if d is not None else ""

def set_cell_text(cell: ET.Element, text: str, ss_type="String"):
    d = cell.find("ss:Data", NS)
    if d is None:
        d = ET.SubElement(cell, f"{{{SS_NS}}}Data")
    d.set(f"{{{SS_NS}}}Type", ss_type)
    d.text = "" if text is None else str(text)

def get_or_create_cell_at_position(row: ET.Element, pos_1b: int) -> ET.Element:
    existing = dict(enumerate_cells_with_positions(row))
    if pos_1b in existing:
        return existing[pos_1b]
    cell = ET.Element(f"{{{SS_NS}}}Cell")
    cell.set(f"{{{SS_NS}}}Index", str(pos_1b))
    row.append(cell)
    return cell

def find_header_row_index(rows: List[ET.Element], anchor="Order ID", scan_limit=80) -> int:
    for i in range(min(scan_limit, len(rows))):
        for _, c in enumerate_cells_with_positions(rows[i]):
            if normalize_cell(get_cell_text(c)) == anchor:
                return i
    raise RuntimeError("Header row not found (anchor='Order ID').")

def header_map(row: ET.Element) -> Dict[str, int]:
    m = {}
    for pos, c in enumerate_cells_with_positions(row):
        t = normalize_cell(get_cell_text(c))
        if t:
            m[t] = pos
    return m

def write_excel_xml(tree: ET.ElementTree, path: str):
    """Write XML and add Excel processing-instruction hint line if missing."""
    buf = io.BytesIO()
    tree.write(buf, encoding="utf-8", xml_declaration=True)
    data = buf.getvalue().decode("utf-8")
    if "<?mso-application" not in data:
        if data.startswith("<?xml"):
            parts = data.split("?>", 1)
            data = parts[0] + "?>\n<?mso-application progid=\"Excel.Sheet\"?>\n" + parts[1]
        else:
            data = "<?mso-application progid=\"Excel.Sheet\"?>\n" + data
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    with open(path, "w", encoding="utf-8", newline="\n") as f:
        f.write(data)

# ========= Optional Excel COM clean-save (Windows) =========

def excel_clean_save(path_in: str, path_out: Optional[str] = None, *, to_xlsx: bool = False) -> Optional[str]:
    """Round-trip the file through Excel to normalize SpreadsheetML.
    Returns the output path if succeeded, else None. Silently no-ops if pywin32/Excel not available.
    """
    try:
        import win32com.client  # type: ignore
        import pythoncom  # type: ignore
    except Exception:
        return None

    file_format = 51 if to_xlsx else 46  # 51=xlsx, 46=Excel 2003 XML
    if path_out is None:
        root, _ = os.path.splitext(path_in)
        path_out = root + (".xlsx" if to_xlsx else ".xml")

    pythoncom.CoInitialize()
    excel = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(os.path.abspath(path_in))
        wb.SaveAs(os.path.abspath(path_out), FileFormat=file_format)
        wb.Close(SaveChanges=False)
        return path_out
    except Exception:
        traceback.print_exc()
        return None
    finally:
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass

# ========= Build a fresh (template-free) workbook with our headers =========
MAIN_HEADERS = [
    "Order ID",
    "Pieces",
    "Pieces Batch",
    "Article Group",
    "Article ID",
    "Wirelist Link",
]

def new_workbook_with_headers(sheet_name="MAIN", extra_cols=0) -> ET.ElementTree:
    # <Workbook>
    wb = ET.Element(f"{{{NAMESPACES['']}}}Workbook")
    # Minimal <Styles> with Default style
    styles = ET.SubElement(wb, f"{{{NAMESPACES['']}}}Styles")
    style = ET.SubElement(styles, f"{{{NAMESPACES['']}}}Style")
    style.set(f"{{{SS_NS}}}ID", "Default")
    style.set(f"{{{SS_NS}}}Name", "Normal")

    # <Worksheet>
    ws = ET.SubElement(wb, f"{{{NAMESPACES['']}}}Worksheet")
    ws.set(f"{{{SS_NS}}}Name", sheet_name)

    # <Table>
    table = ET.SubElement(ws, f"{{{NAMESPACES['']}}}Table")

    # Header row
    hdr_row = ET.SubElement(table, f"{{{NAMESPACES['']}}}Row")
    for i, h in enumerate(MAIN_HEADERS, start=1):
        cell = ET.SubElement(hdr_row, f"{{{NAMESPACES['']}}}Cell")
        cell.set(f"{{{SS_NS}}}Index", str(i))
        set_cell_text(cell, h, "String")

    # Extra headers to the right (template-free mode)
    for j in range(extra_cols):
        i = len(MAIN_HEADERS) + j + 1
        cell = ET.SubElement(hdr_row, f"{{{NAMESPACES['']}}}Cell")
        cell.set(f"{{{SS_NS}}}Index", str(i))
        set_cell_text(cell, f"Extra {j+1}", "String")

    return ET.ElementTree(wb)

# ========= Fill '---' to the right of Wirelist Link =========

def fill_trailing_dashes(hdr: Dict[str, int], row: ET.Element):
    """For any headers positioned to the right of 'Wirelist Link', set cell to '---'."""
    if "Wirelist Link" not in hdr:
        return
    wl_pos = hdr["Wirelist Link"]
    for _, pos in hdr.items():
        if pos > wl_pos:
            cell = get_or_create_cell_at_position(row, pos)
            set_cell_text(cell, "---", "String")

# ========= Read values from a wire XML =========

def read_first_data_row_value(xml_path: str, wanted_header: str, header_anchor="Order ID") -> Optional[str]:
    """Open a wire XML and return the text from the first non-empty data row under `wanted_header`."""
    try:
        t = parse_tree(xml_path)
    except Exception as e:
        print(f"[WARN] Could not parse reference {xml_path}: {e}")
        return None
    root = t.getroot()
    table = find_table(root)
    if table is None:
        return None
    rows = get_rows(table)
    if not rows:
        return None
    try:
        hdr_idx = find_header_row_index(rows, header_anchor)
    except Exception:
        return None
    hdr = header_map(rows[hdr_idx])
    key = normalize_cell(wanted_header)
    if key not in hdr:
        return None
    col = hdr[key]
    # first non-empty data cell in that column
    for r in rows[hdr_idx+1:]:
        pos2cell = dict(enumerate_cells_with_positions(r))
        cell = pos2cell.get(col)
        if cell is None:
            continue
        txt = normalize_cell(get_cell_text(cell))
        if txt != "":
            return txt
    return ""  # nothing non-empty found

# ========= Compose Article ID =========

def compose_article_id(article_group: str, fallback_wirefile_stem: str) -> str:
    """
    Column 5 'Article ID' (operator label).
    Default: equal to Article Group; if empty, fall back to file name stem.
    """
    ag = normalize_cell(article_group)
    return ag if ag else normalize_cell(fallback_wirefile_stem)

# ========= Build a row, return it (so we can fill dashes) =========

def build_main_row(hdr: Dict[str, int], values: Dict[str, Tuple[str, str]]) -> ET.Element:
    """
    Create one <Row/> with given column values.
    values = { header_text: (value, type) }, type in {"String","Number"}
    """
    row = ET.Element(f"{{{NAMESPACES['']}}}Row")
    for header, (val, typ) in values.items():
        hkey = normalize_cell(header)
        if hkey not in hdr:
            continue
        pos = hdr[hkey]
        cell = get_or_create_cell_at_position(row, pos)
        set_cell_text(cell, normalize_cell(val), ss_type=typ)
    return row

# ========= Strip <Table> size attributes (Option B) =========

def strip_table_size_attributes(table: ET.Element):
    """
    Remove attributes that sometimes cause Excel 'Bad Value' errors:
    ss:ExpandedRowCount, ss:ExpandedColumnCount, ss:FullColumns, ss:FullRows
    """
    for attr in ("ExpandedRowCount", "ExpandedColumnCount", "FullColumns", "FullRows"):
        table.attrib.pop(f"{{{SS_NS}}}{attr}", None)

# ========= Build one main for a given list of wire files =========

def build_one_main(template_xml: Optional[str], wire_paths: List[str], out_xml: str,
                   header_anchor="Order ID", extra_cols_after_wl: int = 0,
                   *, do_clean_save: bool = True, clean_to_xlsx: bool = False) -> int:
    # Start from template if provided; otherwise, build workbook fresh.
    if template_xml:
        tpl = parse_tree(template_xml)
        root = tpl.getroot()
        table = find_table(root)
        if table is None:
            raise RuntimeError("Template has no <Table>.")
        rows = get_rows(table)
        hdr_idx = find_header_row_index(rows, anchor=header_anchor)
        hdr = header_map(rows[hdr_idx])
        # clear rows after header
        for r in rows[hdr_idx+1:]:
            table.remove(r)
        tree = tpl
    else:
        tree = new_workbook_with_headers(sheet_name="MAIN", extra_cols=extra_cols_after_wl)
        root = tree.getroot()
        table = find_table(root)
        rows = get_rows(table)
        hdr_idx = 0  # our header is the first row
        hdr = header_map(rows[hdr_idx])

    total = 0
    for ref_path in sorted(wire_paths):
        fname = os.path.basename(ref_path)
        stem = os.path.splitext(fname)[0]  # for Wirelist Link
        # Column 6: Wirelist Link
        wirelist_link = stem
        # Column 4: Article Group (from wire file)
        ag = read_first_data_row_value(ref_path, "Article Group", header_anchor=header_anchor)
        if ag is None:
            ag = ""
        ag = normalize_cell(ag)
        # Column 5: Article ID (operator label)
        article_id = compose_article_id(ag, stem)
        # If the Wirelist Link has a chunk suffix like _01 or _02, mirror it in the Article ID
        m_sfx = re.search(r"_(\d{2})$", stem)
        row_suffix = m_sfx.group(0) if m_sfx else ""
        if row_suffix and not article_id.endswith(row_suffix):
            article_id = f"{article_id}{row_suffix}"

        row_vals = {
            "Order ID":        ("1", "Number"),
            "Pieces":          ("1", "Number"),
            "Pieces Batch":    ("1", "Number"),
            "Article Group":   (ag, "String"),
            "Article ID":      (article_id, "String"),
            "Wirelist Link":   (wirelist_link, "String"),
        }
        row = build_main_row(hdr, row_vals)
        # fill '---' to the right of Wirelist Link (works with template + template-free extra headers)
        fill_trailing_dashes(hdr, row)
        table.append(row)
        total += 1

    strip_table_size_attributes(table)
    write_excel_xml(tree, out_xml)

    # Excel COM clean-save (optional)
    if do_clean_save:
        excel_clean_save(out_xml, None, to_xlsx=clean_to_xlsx)

    return total

# ========= Build mains per Section (combine all panels) =========

def build_mains_per_section(template_xml: Optional[str], wires_dir: str, outdir: str,
                            header_anchor="Order ID", extra_cols_after_wl: int = 0,
                            *, do_clean_save: bool = True, clean_to_xlsx: bool = False):
    if not os.path.isdir(wires_dir):
        raise RuntimeError(f"Wires directory not found: {wires_dir}")

    wire_files = [os.path.join(wires_dir, f) for f in os.listdir(wires_dir) if f.lower().endswith(".xml")]
    if not wire_files:
        raise RuntimeError(f"No .xml files found in {wires_dir}")

    # group files by Gauge+Color (e.g., 14WHT, 14GRY)
    groups: DefaultDict[str, List[str]] = defaultdict(list)
    for path in wire_files:
        stem = os.path.splitext(os.path.basename(path))[0]
        gauge_color = parse_gauge_color_from_stem(stem)
        groups[gauge_color].append(path)

    os.makedirs(outdir, exist_ok=True)

    totals = {}
    for gauge_color, paths in sorted(groups.items()):
        paths = sorted(paths)
        out_xml_base = os.path.join(outdir, f"{gauge_color}_main")
        out_xml = out_xml_base + (".xlsx" if clean_to_xlsx else ".xml")
        count = build_one_main(template_xml, paths, out_xml,
                               header_anchor=header_anchor,
                               extra_cols_after_wl=extra_cols_after_wl,
                               do_clean_save=do_clean_save,
                               clean_to_xlsx=clean_to_xlsx)
        totals[gauge_color] = count
        print(f"Wrote {count} row(s) -> {out_xml}")

    print(f"Built {len(totals)} MAIN file(s).")
    return totals

# ========= CLI =========

def main():
    ap = argparse.ArgumentParser(
        description="Build MAIN SpreadsheetMLs per Section (combine all panels). Template optional; fills trailing columns with ---. Includes optional Excel clean-save."
    )
    ap.add_argument("--template", required=False, default=None,
                    help="Main template XML (optional; keeps styles if provided)")
    ap.add_argument("--wires-dir", required=True, help="Folder with per-wire XMLs")
    ap.add_argument("--outdir", required=True, help="Output folder for per-section MAIN XMLs")
    ap.add_argument("--header-anchor", default="Order ID",
                    help="Header text to find the header row (only used if template provided)")
    ap.add_argument("--fill-after", type=int, default=0,
                    help="Template-free mode: create N extra columns AFTER 'Wirelist Link' and fill them with '---'")
    ap.add_argument("--no-clean-save", action="store_true", help="Skip Excel clean-save roundtrip (default: enabled)")
    ap.add_argument("--xlsx", action="store_true", help="Save cleaned files as .xlsx instead of .xml")
    args = ap.parse_args()

    build_mains_per_section(args.template, args.wires_dir, args.outdir,
                            header_anchor=args.header_anchor,
                            extra_cols_after_wl=args.fill_after,
                            do_clean_save=(not args.no_clean_save),
                            clean_to_xlsx=args.xlsx)

if __name__ == "__main__":
    main()
