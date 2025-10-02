#!/usr/bin/env python3
import argparse, os, re, copy, math, io, traceback
import xml.etree.ElementTree as ET
import json
from typing import List, Tuple, Optional, Dict

# =========================
# Namespaces (Excel 2003 XML)
# =========================
NAMESPACES = {
    "":     "urn:schemas-microsoft-com:office:spreadsheet",
    "o":    "urn:schemas-microsoft-com:office:office",
    "x":    "urn:schemas-microsoft-com:office:excel",
    "ss":   "urn:schemas-microsoft-com:office:spreadsheet",
    "html": "http://www.w3.org/TR/REC-html40",
}
for prefix, uri in NAMESPACES.items():
    ET.register_namespace(prefix, uri)

SPREADSHEET_NS = NAMESPACES[""]
SS_NS = NAMESPACES["ss"]
NS = {"ss": SS_NS}

# =========================
# Helpers
# =========================
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
            if get_cell_text(c).strip() == anchor:
                return i
    raise RuntimeError("Header row not found (anchor='Order ID').")

def header_map(row: ET.Element) -> Dict[str, int]:
    m = {}
    for pos, c in enumerate_cells_with_positions(row):
        t = get_cell_text(c).strip()
        if t:
            m[t] = pos
    return m

def write_excel_xml(tree: ET.ElementTree, path: str):
    """Write XML and (if missing) add the Excel processing-instruction hint line."""
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

def _has_plcio(text: str) -> bool:
    return "PLCIO" in (text or "").upper()
        

# =========================
# Optional Excel COM clean-save (Windows)
# =========================
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

# =========================
# Regexes for parsing
# =========================
_wireid_re  = re.compile(r"^\s*(\d+)\s*-\s*([A-Za-z]+)\s*$")  # 18-WHT → 18, WHT
_section_re = re.compile(r"[sS]\d+[A-Za-z]?")                 # e.g., s5D
_panel_re   = re.compile(r"[pP][A-Za-z0-9]+")                 # e.g., pC14
_job_re     = re.compile(r"^\s*([A-Za-z0-9]+)")

def parse_section_panel_from_filename(fname: str) -> Tuple[str, str]:
    base = os.path.splitext(os.path.basename(fname))[0]
    s = _section_re.search(base)
    p = _panel_re.search(base)
    return (s.group(0) if s else "null", p.group(0) if p else "null")

def parse_gc(text: str) -> str:
    m = _wireid_re.match(text.strip())
    return f"{m.group(1)}{m.group(2)}" if m else "null"

# =========================
# Fixers
# =========================
def clear_printer_texts(tree: ET.ElementTree, header_anchor="Order ID") -> int:
    """Clear 'Printer1 Wire1 BeginText' and 'Printer1 Wire1 EndText' across all data rows."""
    root = tree.getroot(); table = find_table(root)
    rows = get_rows(table); hdr_idx = find_header_row_index(rows, header_anchor); hdr = header_map(rows[hdr_idx])
    changed = 0
    for name in ("Printer1 Wire1 BeginText", "Printer1 Wire1 EndText"):
        if name not in hdr:
            print(f"[WARN] Header not found: {name!r}")
            continue
        col = hdr[name]
        for r in rows[hdr_idx + 1:]:
            c = dict(enumerate_cells_with_positions(r)).get(col)
            if c is None:
                continue
            if get_cell_text(c):
                set_cell_text(c, "", "String"); changed += 1
    return changed

def set_printer1_from_wireid(tree: ET.ElementTree, header_anchor="Order ID") -> int:
    """Set 'Printer1 ID' = 'AWG ' + WireID (hyphens/spaces removed)."""
    root = tree.getroot(); table = find_table(root)
    rows = get_rows(table); hdr_idx = find_header_row_index(rows, header_anchor); hdr = header_map(rows[hdr_idx])
    if "Wire ID" not in hdr or "Printer1 ID" not in hdr:
        print("[WARN] Missing 'Wire ID' or 'Printer1 ID' headers.")
        return 0
    col_wire = hdr["Wire ID"]; col_prn = hdr["Printer1 ID"]; changed = 0
    for r in rows[hdr_idx + 1:]:
        wire_cell = dict(enumerate_cells_with_positions(r)).get(col_wire)
        wire_text = (get_cell_text(wire_cell).strip() if wire_cell is not None else "")
        if not wire_text:
            continue
        cleaned = wire_text.replace("-", "").replace(" ", "")
        new_val = f"AWG {cleaned}"
        prn_cell = get_or_create_cell_at_position(r, col_prn)
        if get_cell_text(prn_cell) != new_val:
            set_cell_text(prn_cell, new_val, "String"); changed += 1
    return changed

def set_article_group(tree: ET.ElementTree, source_path: str, header_anchor="Order ID") -> int:
    """
    'Article Group' = '<job> <Section><Panel><Gauge><Color>' or '<job> null'
    job: first token already present in Article Group cell, else guessed from filename (e.g., 20321P)
    Section/Panel: parsed from filename (preserve case as found)
    Gauge/Color: from Wire ID (e.g., 18-WHT → 18WHT)
    """
    root = tree.getroot(); table = find_table(root)
    rows = get_rows(table); hdr_idx = find_header_row_index(rows, header_anchor); hdr = header_map(rows[hdr_idx])
    needed = ("Article Group", "Wire ID")
    for n in needed:
        if n not in hdr:
            print(f"[WARN] Missing header: {n!r}")
            return 0
    col_ag = hdr["Article Group"]; col_wire = hdr["Wire ID"]
    section, panel = parse_section_panel_from_filename(source_path)

    changed = 0
    base = os.path.splitext(os.path.basename(source_path))[0]
    for r in rows[hdr_idx + 1:]:
        pos2cell = dict(enumerate_cells_with_positions(r))
        ag_cell = pos2cell.get(col_ag); cur_ag = get_cell_text(ag_cell) if ag_cell is not None else ""
        # job from current AG or fallback to filename token like 5+ digits + optional letter
        job = None
        m_job = _job_re.match(cur_ag or "")
        if m_job:
            job = m_job.group(1)
        else:
            m_name = re.search(r"([0-9]{4,}[A-Za-z]?)", base)
            if m_name:
                job = m_name.group(1)
        # gauge+color
        wire_cell = pos2cell.get(col_wire)
        gc = parse_gc(get_cell_text(wire_cell) if wire_cell is not None else "")
        suffix = f"{section}{panel}{gc}" if (section != "null" and panel != "null" and gc != "null") else "null"
        final = f"{job or 'null'} {suffix}"
        if ag_cell is None:
            ag_cell = get_or_create_cell_at_position(r, col_ag)
        if get_cell_text(ag_cell) != final:
            set_cell_text(ag_cell, final, "String"); changed += 1
    return changed

def set_last_three_distances(tree: ET.ElementTree, header_anchor="Order ID") -> int:
    """Set the distance columns to 0.79, 4, 0.79 (Number) for every data row."""
    root = tree.getroot(); table = find_table(root); rows = get_rows(table)
    hdr_idx = find_header_row_index(rows, header_anchor); hdr = header_map(rows[hdr_idx])
    needed = ["Printer1 Wire1 BeginDistance", "Printer1 Wire1 EndlessDistance", "Printer1 Wire1 EndDistance"]
    for n in needed:
        if n not in hdr:
            print(f"[WARN] Missing header: {n!r}")
            return 0
    cols = [hdr[needed[0]], hdr[needed[1]], hdr[needed[2]]]
    vals = [0.79, 4, 0.79]; changed = 0
    for r in rows[hdr_idx + 1:]:
        for col, v in zip(cols, vals):
            c = get_or_create_cell_at_position(r, col)
            if get_cell_text(c) != str(v):
                set_cell_text(c, v, "Number"); changed += 1
    return changed

# Rule for null files and wire length
def fix_wire_length_for_null_files(tree: ET.ElementTree, source_path: str, header_anchor="Order ID") -> int:
    """If the filename contains 'null', any row with Wire Length=300 becomes 200."""
    if "null" not in os.path.basename(source_path):
        return 0
    root = tree.getroot(); table = find_table(root); rows = get_rows(table)
    hdr_idx = find_header_row_index(rows, header_anchor); hdr = header_map(rows[hdr_idx])
    if "Wire Length" not in hdr:
        print("[WARN] Missing 'Wire Length' header; cannot apply null-length fix.")
        return 0
    col_len = hdr["Wire Length"]; changed = 0
    for r in rows[hdr_idx + 1:]:
        c = dict(enumerate_cells_with_positions(r)).get(col_len)
        if c is None:
            continue
        val = get_cell_text(c).strip()
        if val == "300":
            set_cell_text(c, "200", "Number"); changed += 1
    return changed

# =========================
# Auto-Crimp (default ON) — updated matching & side selection
# =========================
# - Plans to add more ID tokens. 
# - Plans to change logic as to which colunm to assign Crimp ID to. 
# - Plans to be able to apply different crimp ID's. 
# - Plans to be able for this to work on other panels and with other

# (Sets retained for reference; detection uses regex below.)
TRIGGER_TOKENS_FIXED = {"PSS1","PSS2","PSS3","PSS4"}
TRIGGER_TOKENS_AB    = {"VMSS","AMSS","CBCS","86","WM","FM"}

# Accept:
#  - PSS1..PSS4
#  - (VMSS|AMSS|WM|FM|86) with optional A/B (e.g., WMA, FMB, 86B)
#  - CBCS with ANY alphanumeric prefix (e.g., F4CBCS, TCBCS) and optional A/B
ARTICLE_TOKEN_RE = re.compile(
    r"\b(?:PSS[1-4]|(?:VMSS|AMSS|WM|FM|86)(?:[AB])?|[A-Z0-9]*CBCS(?:[AB])?)\b",
    re.IGNORECASE
)

def _parse_panel_gauge_from_p_token(p_token: str) -> Tuple[Optional[str], Optional[int]]:
    """From something like 'pC14' return ('C', 14)."""
    if not p_token or len(p_token) < 2:
        return (None, None)
    m = re.match(r"[pP]([A-Za-z])(\d{1,2})$", p_token)
    if not m:
        return (None, None)
    try:
        return (m.group(1).upper(), int(m.group(2)))
    except Exception:
        return (m.group(1).upper(), None)

def _extract_gauge_from_wire_id(text: str) -> Optional[int]:
    if not text:
        return None
    m = re.search(r"\b(\d{1,2})\b", text)
    if not m:
        return None
    try:
        return int(m.group(1))
    except:
        return None

def _first_last_endpoint_tokens(aid: str) -> Tuple[Optional[str], Optional[str]]:
    """Split Article ID on whitespace; take first and last tokens; strip text after ':'."""
    if not aid:
        return None, None
    parts = [p for p in re.split(r"\s+", aid.strip()) if p]
    if not parts:
        return None, None
    left = parts[0].split(":", 1)[0]
    right = parts[-1].split(":", 1)[0]
    return (left, right)

def _token_matches_endpoint(tok: Optional[str]) -> bool:
    return bool(tok and ARTICLE_TOKEN_RE.fullmatch(tok.upper()))

def load_rules_file(path: str) -> Optional[dict]:
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        # compile regexes
        for rule in data.get("rules", []):
            for key in ("tokens_left", "tokens_right", "tokens_any"):
                pats = rule.get(key, [])
                rule[key] = [re.compile(p, re.IGNORECASE) for p in pats]
        return data
    except Exception as e:
        print(f"[WARN] Failed to load rules file {path}: {e}")
        return None

def _rule_matches_panel_gauge(rule: dict, panel_letter: Optional[str],
                              filename_gauge: Optional[int], wire_id_text: str) -> bool:
    panels = rule.get("panels")
    if panels and panel_letter and panel_letter.upper() not in [p.upper() for p in panels]:
        return False

    gauges = rule.get("gauges")
    if gauges:
        ok = False
        if filename_gauge in gauges:
            ok = True
        else:
            g = _extract_gauge_from_wire_id(wire_id_text or "")
            if g in gauges:
                ok = True
        if not ok:
            return False
    return True

def _decide_side_by_rule(rule: dict, ltok: Optional[str], rtok: Optional[str], prefer: str) -> Optional[int]:
    def any_match(tok: Optional[str], regexes: List[re.Pattern]) -> bool:
        return bool(tok and any(r.search(tok) for r in regexes))

    left  = any_match(ltok, rule.get("tokens_left", []))
    right = any_match(rtok, rule.get("tokens_right", []))

    if not left and not right:
        # optional catch-all if you want “either side” triggers
        any_side = rule.get("tokens_any", [])
        if any_side:
            left  = any_match(ltok, any_side)
            right = any_match(rtok, any_side)

    if left and right:
        return 15 if prefer == "left" else 19
    if left:
        return 15
    if right:
        return 19
    return None

#TESTING USING JSON RULES
def apply_crimp_rules(tree: ET.ElementTree, source_path: str, rules: dict, *, header_anchor="Order ID") -> int:
    root = tree.getroot(); table = find_table(root)
    rows = get_rows(table); hdr_idx = find_header_row_index(rows, header_anchor); hdr = header_map(rows[hdr_idx])

    col_wire = hdr.get("Wire ID", 11)
    col_aid  = hdr.get("Article ID", 6)

    # filename panel/gauge
    _, p_token = parse_section_panel_from_filename(source_path)
    panel_letter, filename_gauge = _parse_panel_gauge_from_p_token(p_token)

    prefer_default = (rules.get("defaults", {}) or {}).get("prefer_when_both", "left")
    changed = 0

    for r in rows[hdr_idx + 1:]:
        pos2cell = dict(enumerate_cells_with_positions(r))
        wire_txt = get_cell_text(pos2cell.get(col_wire)) if col_wire in pos2cell else ""
        aid      = (get_cell_text(pos2cell.get(col_aid)) if col_aid in pos2cell else "").strip()

        if not aid:
            continue
        ltok, rtok = _first_last_endpoint_tokens(aid)

        # try rules in order
        for rule in rules.get("rules", []):
            if not _rule_matches_panel_gauge(rule, panel_letter, filename_gauge, wire_txt):
                continue

            prefer = (rule.get("prefer_when_both") or prefer_default).lower()
            target = _decide_side_by_rule(rule, ltok, rtok, prefer)
            if target not in (15, 19):
                continue

            # columns (allow override per rule; default 15/19)
            cols = rule.get("columns", {}) or {}
            left_col  = int(cols.get("left", 15))
            right_col = int(cols.get("right", 19))
            target_col = left_col if target == 15 else right_col

            # current values / protect non-overwrite
            v_left  = get_cell_text(get_or_create_cell_at_position(r, left_col))
            v_right = get_cell_text(get_or_create_cell_at_position(r, right_col))
            if v_left.strip() and v_right.strip():
                break  # next row
            if v_left.strip() and target_col == left_col and v_left.strip() != rule["crimp_id"]:
                break
            if v_right.strip() and target_col == right_col and v_right.strip() != rule["crimp_id"]:
                break

            # if chosen side occupied but the other is free, flip
            if target_col == left_col and v_left.strip() and not v_right.strip():
                target_col = right_col
            elif target_col == right_col and v_right.strip() and not v_left.strip():
                target_col = left_col

            tgt = get_or_create_cell_at_position(r, target_col)
            if not get_cell_text(tgt).strip():
                set_cell_text(tgt, rule["crimp_id"], "String")
                # blank two left neighbors of whichever side we actually wrote
                if target_col == left_col:
                    for k in (left_col - 1, left_col - 2):
                        set_cell_text(get_or_create_cell_at_position(r, k), "", "String")
                else:
                    for k in (right_col - 1, right_col - 2):
                        set_cell_text(get_or_create_cell_at_position(r, k), "", "String")
                changed += 1
            break  # stop at first matching rule for this row

    return changed


def apply_auto_crimp_endpoints(
    tree: ET.ElementTree,
    source_path: str,
    *,
    header_anchor="Order ID",
    crimp_id="018769-025",
    prefer_when_both="left"
) -> int:
    """
    - Applies to files that are Panel C + 14 AWG (from filename 'pC14' and/or Wire ID col 11).
    - Parses Article ID and evaluates the FIRST and LAST tokens (before ':'):
        * If left token matches → write to col 15 (left/side 0)
        * Else if right token matches → write to col 19 (right/side 1)
        * If both match → use prefer_when_both ('left' or 'right')
    - Never both sides. Never overwrite a different existing value.
    - When writing, blanks the two immediate left neighbors of that side.
    """
    root = tree.getroot(); table = find_table(root)
    rows = get_rows(table); hdr_idx = find_header_row_index(rows, header_anchor); hdr = header_map(rows[hdr_idx])

    # Columns used
    col_wire = hdr.get("Wire ID", 11)     # fallback
    col_aid  = hdr.get("Article ID", 6)   # fallback
    col15, col19 = 15, 19

    # File-based panel/gauge
    section_token, panel_token = parse_section_panel_from_filename(source_path)
    panel_letter, gauge_from_name = _parse_panel_gauge_from_p_token(panel_token)

    # If file name clearly indicates not C or not 14, skip entirely
    if panel_letter and panel_letter != "C":
        return 0
    if gauge_from_name and gauge_from_name != 14:
        return 0

    changed = 0
    for r in rows[hdr_idx + 1:]:
        pos2cell = dict(enumerate_cells_with_positions(r))

        # Gauge gate: allow either filename gauge=14 OR Wire ID contains 14
        gauge_ok = False
        if gauge_from_name == 14:
            gauge_ok = True
        else:
            wtxt = get_cell_text(pos2cell.get(col_wire)) if col_wire in pos2cell else ""
            g = _extract_gauge_from_wire_id(wtxt)
            if g == 14:
                gauge_ok = True
        if not gauge_ok:
            continue

        # Panel gate: if filename gave panel, require C; if filename didn't, allow pass-through
        if panel_letter and panel_letter != "C":
            continue

        # Article ID
        aid = get_cell_text(pos2cell.get(col_aid)) if col_aid in pos2cell else ""
        aid = (aid or "").strip()
        if not aid:
            continue

        # Decide side using first/last tokens only
        ltok, rtok = _first_last_endpoint_tokens(aid)
        left_ok = _token_matches_endpoint(ltok)
        right_ok = _token_matches_endpoint(rtok)
        if not left_ok and not right_ok:
            continue

        # Current values
        v15 = get_cell_text(pos2cell.get(col15)) if col15 in pos2cell else ""
        v19 = get_cell_text(pos2cell.get(col19)) if col19 in pos2cell else ""
        if v15.strip() and v19.strip():
            continue  # both already set
        if v15.strip() and v15.strip() != crimp_id:
            continue  # don't overwrite different value
        if v19.strip() and v19.strip() != crimp_id:
            continue

        # Choose target
        prefer_col = 15 if prefer_when_both.lower() == "left" else 19
        if left_ok and right_ok:
            target = prefer_col
        elif left_ok:
            target = 15
        else:
            target = 19

        # If chosen is occupied but the other side is empty, flip
        if target == 15 and v15.strip() and not v19.strip():
            target = 19
        elif target == 19 and v19.strip() and not v15.strip():
            target = 15

        # Write if empty and blank the two left neighbors
        tgt_cell = get_or_create_cell_at_position(r, target)
        if not get_cell_text(tgt_cell).strip():
            set_cell_text(tgt_cell, crimp_id, "String")
            if target == 15:
                for k in (14, 13):  # neighbors of col 15
                    set_cell_text(get_or_create_cell_at_position(r, k), "", "String")
            else:
                for k in (18, 17):  # neighbors of col 19
                    set_cell_text(get_or_create_cell_at_position(r, k), "", "String")
            changed += 1

    return changed

# =========================
# Splitter (gauge+color, chunk at 150)
# =========================
def clone_with_header(tree: ET.ElementTree, hdr_idx: int) -> ET.ElementTree:
    nt = copy.deepcopy(tree)
    tab = find_table(nt.getroot()); rs = get_rows(tab)
    for r in rs[hdr_idx + 1:]:
        tab.remove(r)
    return nt

def split_by_gauge_color(
    tree: ET.ElementTree,
    src: str,
    outdir: str,
    header_anchor="Order ID",
    max_per=150,
    *,
    do_clean_save: bool = True,
    clean_to_xlsx: bool = False
) -> List[str]:
    """
    Split rows by Gauge+Color (from Wire ID), but also separate any rows whose Article ID
    contains 'PLCIO' (case-insensitive) into their own subgroup within the same Gauge+Color.
    Output files for PLCIO subgroups are suffixed with '_PLCIO'.
    """
    root = tree.getroot(); table = find_table(root); rows = get_rows(table)
    hdr_idx = find_header_row_index(rows, header_anchor); hdr = header_map(rows[hdr_idx])

    if "Wire ID" not in hdr:
        print("[WARN] Missing 'Wire ID' header; cannot split.")
        return []

    col_wire = hdr["Wire ID"]
    col_aid = hdr.get("Article ID")  # may be None if header missing

    # group: key = (gc, is_plcio_bool)
    groups: Dict[Tuple[str, bool], List[ET.Element]] = {}

    for r in rows[hdr_idx + 1:]:
        pos2cell = dict(enumerate_cells_with_positions(r))

        # Gauge+Color
        wire = get_cell_text(pos2cell.get(col_wire)) if col_wire in pos2cell else ""
        gc = parse_gc(wire)  # e.g., 18-WHT -> 18WHT, else "null"

        # PLCIO flag (only if Article ID column exists)
        is_plcio = False
        if col_aid is not None:
            aid_txt = get_cell_text(pos2cell.get(col_aid)) if col_aid in pos2cell else ""
            is_plcio = _has_plcio(aid_txt)
        else:
            # If there's no Article ID, we cannot detect PLCIO; keep legacy behavior
            pass

        groups.setdefault((gc, is_plcio), []).append(r)

    # naming prefix from filename
    section, panel = parse_section_panel_from_filename(src)
    os.makedirs(outdir, exist_ok=True)

    written: List[str] = []
    for (gc, is_plcio), rs in groups.items():
        key_base = f"{section}{panel}{gc}"
        key = key_base + ("_PLCIO" if is_plcio else "")
        total = len(rs)
        chunks = max(1, math.ceil(total / max_per))

        for i in range(chunks):
            chunk = rs[i * max_per:(i + 1) * max_per]

            nt = clone_with_header(tree, hdr_idx)
            tab = find_table(nt.getroot())
            for r in chunk:
                tab.append(copy.deepcopy(r))

            name = key + (f"_{i+1:02d}" if chunks > 1 else "")
            out_xml = os.path.join(outdir, f"{name}.xml")
            write_excel_xml(nt, out_xml)

            cleaned_path = None
            if do_clean_save:
                cleaned_path = excel_clean_save(out_xml, None, to_xlsx=clean_to_xlsx)
            written.append(cleaned_path or out_xml)

    # If Article ID was missing, notify once (we already split without PLCIO separation)
    if "Article ID" not in hdr:
        print("[WARN] 'Article ID' header not found; PLCIO separation was skipped.")

    return written


# =========================
# CLI
# =========================
def main():
    ap = argparse.ArgumentParser(
        description="Fix and split Excel 2003 XML (SpreadsheetML) with optional Excel clean-save."
    )
    ap.add_argument("--in", dest="in_paths", required=True, nargs="+", help="Input XML file(s)")
    ap.add_argument("--outdir", required=True, help="Output folder for split files")
    ap.add_argument("--header-anchor", default="Order ID", help="Header row anchor text (default: 'Order ID')")
    ap.add_argument("--max-per-file", type=int, default=150, help="Max wires per output file (default 150)")
    ap.add_argument("--no-clean-save", action="store_true", help="Skip Excel clean-save roundtrip (default: enabled)")
    ap.add_argument("--xlsx", action="store_true", help="Save cleaned files as .xlsx instead of .xml")

    # Auto-crimp (DEFAULT ON; use --no-auto-crimp to disable)
    ap.add_argument(
        "--no-auto-crimp",
        action="store_true",
        help="Disable Auto Crimp. By default, assigns crimp 018769-025 for Panel C + 14 AWG rows based on endpoints in Article ID."
    )
    ap.add_argument("--crimp-id", default="018769-025", help="Crimp ID to assign (default: 018769-025).")
    ap.add_argument("--prefer-when-both", choices=["left", "right"], default="left",
                    help="If both ends match, prefer left→col 15 or right→col 19 (default: left).")
    ap.add_argument("--rules", help="Path to crimp rules JSON. If omitted, auto-load ./crimp_rules.json when present.")


    args = ap.parse_args()

    for in_path in args.in_paths:
        tree = parse_tree(in_path) 
        # Auto-load rules file if present next to the script
        rules_path = args.rules
        if not rules_path:
            default_rules = os.path.join(os.path.dirname(__file__), "crimp_rules.json")
            if os.path.exists(default_rules):
                rules_path = default_rules

        rules = load_rules_file(rules_path) if rules_path else None
 
    

    

        # ---- Apply ALL requested fixes (in-place) ----
        n1 = clear_printer_texts(tree, header_anchor=args.header_anchor)
        n2 = set_printer1_from_wireid(tree, header_anchor=args.header_anchor)
        n3 = set_article_group(tree, source_path=in_path, header_anchor=args.header_anchor)
        n4 = set_last_three_distances(tree, header_anchor=args.header_anchor)
        n5 = fix_wire_length_for_null_files(tree, source_path=in_path, header_anchor=args.header_anchor)

        # Auto-crimp (before splitting)
        n6 = 0
        if not args.no_auto_crimp:
            if rules and rules.get("rules"):
                n6 = apply_crimp_rules(tree, in_path, rules, header_anchor=args.header_anchor)
            else:
            # Fallback to current hardcoded C14 logic
                n6 = apply_auto_crimp_endpoints(
                    tree,
                    source_path=in_path,
                    header_anchor=args.header_anchor,
                    crimp_id=args.crimp_id,
                    prefer_when_both=args.prefer_when_both
            )

        print(
            f"Cleared Begin/EndText: {n1} | Set Printer1 ID: {n2} | Set Article Group: {n3} | "
            f"Distances set: {n4} | Wire length fixes (300→200 on null files): {n5} | Auto Crimp Set: {n6}"
        )

        # ---- Split, write, and optionally clean-save outputs ----
        outs = split_by_gauge_color(
            tree,
            src=in_path,
            outdir=args.outdir,
            header_anchor=args.header_anchor,
            max_per=args.max_per_file,
            do_clean_save=(not args.no_clean_save),
            clean_to_xlsx=args.xlsx,
        )
        print(f"Created {len(outs)} file(s):")
        for o in outs:
            print(" -", o)

if __name__ == "__main__":
    main()
