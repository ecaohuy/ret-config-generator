"""Core CDD -> RETConfigWDTInternal conversion logic (no UI).

Shared by generate_ret.py (CLI) and ret_gui.py (Tkinter GUI).
"""
import json
from collections import defaultdict
from copy import copy

from openpyxl import load_workbook

# 4 devices per sector: order matches positions 1..4 in the template.
DEVICE_ORDER = ["C", "D", "E", "_NOT_USED_"]

# Output column headers (target order).
HEADERS = [
    "Site Name",
    "RRU Name",
    "RRU CN",
    "RRU SRN",
    "RRU SN",
    "RCU Coloring",
    "RCU Tilt",
    "Device Name",
]


def load_mapping(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def _norm_header(s):
    """Normalize a header cell for tolerant matching (case/whitespace-insensitive)."""
    if s is None:
        return ""
    return " ".join(str(s).split()).strip().lower()


def _col_letter_to_index(letter):
    """'A' -> 0, 'B' -> 1, ... (single-letter columns)."""
    return ord(str(letter).strip().upper()[0]) - ord("A")


def _find_index(norm_headers, targets):
    """Index of the first header matching any target (exact, then startswith)."""
    targets = [t for t in targets if t]
    for t in targets:
        if t in norm_headers:
            return norm_headers.index(t)
    for i, h in enumerate(norm_headers):
        if h and any(h.startswith(t) for t in targets):
            return i
    return None


def _locate_header(all_rows, key_targets, max_scan=10):
    """Find the header row by locating the key column; return (row_index, norm_headers).

    Scans the first ``max_scan`` rows so a leading title/blank row doesn't break
    detection. Falls back to row 0 if the key column is never found.
    """
    for i, row in enumerate(all_rows[:max_scan]):
        norm = [_norm_header(h) for h in row]
        if _find_index(norm, key_targets) is not None:
            return i, norm
    return 0, [_norm_header(h) for h in (all_rows[0] if all_rows else ())]


def list_sheets(cdd_path):
    """Return sheet names in the CDD workbook."""
    wb = load_workbook(cdd_path, read_only=True)
    try:
        return wb.sheetnames
    finally:
        wb.close()


def load_three_g_tilts(cdd_wb, mapping):
    """Map (site_old, sector_number) -> New E-Tilt from the 3G sheet.

    The 3G Installation Design sheet lists a 'Logical Sector Name' formatted as
    '{SiteName_Old}-{sector number}' (e.g. 'KGPQ01-2'). Device No 3 (band E) is
    emitted as NSN_U2100 only when its sector appears here (else NOT_USED), and
    its tilt is taken from the 'New E-Tilt' column. A sector can span multiple
    rows (carriers); the first non-blank New E-Tilt wins. Membership in the dict
    means the sector is present. Empty dict if the sheet/config is absent.
    """
    cfg = mapping.get("three_g")
    if not cfg or cfg["sheet"] not in cdd_wb.sheetnames:
        return {}
    ws = cdd_wb[cfg["sheet"]]
    all_rows = list(ws.iter_rows(min_row=1, values_only=True))
    if not all_rows:
        return {}

    # Locate columns by header name. The site MATCH key comes from the 'Site'
    # column (== 4G SiteName_Old); only the sector number comes from the
    # Logical Sector Name (text after the last separator).
    site_targets = [_norm_header(cfg.get("site_header", "Site"))]
    sec_targets = [_norm_header(cfg.get("logical_sector_header", "")),
                   _norm_header(cfg.get("logical_sector_header_prefix", "Logical Sector Name"))]
    tilt_targets = [_norm_header(cfg.get("etilt_header", "New E-Tilt"))]
    hdr_idx, norm = _locate_header(all_rows, sec_targets)
    site_col = _find_index(norm, site_targets)
    sec_col = _find_index(norm, sec_targets)
    tilt_col = _find_index(norm, tilt_targets)
    # Fall back to configured column letters if a header isn't found.
    if site_col is None and cfg.get("site_column"):
        site_col = _col_letter_to_index(cfg["site_column"])
    if sec_col is None and cfg.get("logical_sector_column"):
        sec_col = _col_letter_to_index(cfg["logical_sector_column"])
    if tilt_col is None and cfg.get("etilt_column"):
        tilt_col = _col_letter_to_index(cfg["etilt_column"])
    if sec_col is None:
        return {}

    sep = cfg.get("separator", "-")
    tilts = {}
    for r in all_rows[hdr_idx + 1:]:
        if not r or len(r) <= sec_col or r[sec_col] is None:
            continue
        # Sector number = last segment of the Logical Sector Name (rpartition
        # so a site name containing the separator doesn't confuse it).
        _, _, num = str(r[sec_col]).rpartition(sep)
        if not num.strip():
            continue
        # Site comes from the 'Site' column when available, else the LSN prefix.
        if site_col is not None and len(r) > site_col and r[site_col] is not None:
            site_old = str(r[site_col])
        else:
            site_old = str(r[sec_col]).rpartition(sep)[0]
        key = (site_old.strip(), num.strip())
        e_tilt = r[tilt_col] if (tilt_col is not None and len(r) > tilt_col) else None
        try:
            e_tilt_val = float(e_tilt) if e_tilt not in (None, "-", "") else None
        except (TypeError, ValueError):
            e_tilt_val = None
        # First non-blank tilt wins; still register the key if only blanks seen.
        if key not in tilts or (tilts[key] is None and e_tilt_val is not None):
            tilts[key] = e_tilt_val
    return tilts


def build_rows(cdd_path, sheet, mapping):
    """Parse the CDD sheet and produce output rows.

    Returns (rows, skipped, sector_count) where:
      rows      = list of 8-value lists (matching HEADERS)
      skipped   = list of (cell_name, reason)
      sector_count = number of distinct sectors emitted
    """
    band_map = mapping["band_map"]
    sector_map = mapping["sector_map"]
    rru_cn = mapping["constants"]["rru_cn"]
    rru_sn = mapping["constants"]["rru_sn"]
    three_g = mapping.get("three_g", {})
    tg_band = three_g.get("governs_band")            # e.g. "E"
    tg_absent = three_g.get("absent_band_token", "NOT_USED")

    cdd_wb = load_workbook(cdd_path, data_only=True)
    cdd_ws = cdd_wb[sheet]
    all_rows = list(cdd_ws.iter_rows(min_row=1, values_only=True))
    three_g_tilts = load_three_g_tilts(cdd_wb, mapping)
    cdd_wb.close()

    # Resolve columns by HEADER NAME so reordered/extra columns don't break parsing.
    hdr_cfg = mapping.get("source", {}).get("headers", {})
    wanted = {
        "site_old": hdr_cfg.get("site_old", "SiteName (RRU Location)_Old"),
        "site_new": hdr_cfg.get("site_new", "SiteName (RRU Location)_New"),
        "ne_id": hdr_cfg.get("ne_id", "Ne ID (New)"),
        "cell_name": hdr_cfg.get("cell_name", "CellName (New)[Key]"),
        "e_tilt": hdr_cfg.get("e_tilt", "E_TILT"),
    }
    cell_targets = [_norm_header(wanted["cell_name"])]
    hdr_idx, norm = _locate_header(all_rows, cell_targets)
    col = {f: _find_index(norm, [_norm_header(name)]) for f, name in wanted.items()}

    required = ("site_old", "site_new", "ne_id", "cell_name")
    missing = [wanted[f] for f in required if col[f] is None]
    if missing:
        found = [str(h) for h in (all_rows[hdr_idx] if all_rows else ()) if h is not None]
        raise ValueError(
            "Sheet '%s': could not find required column(s) by header name: %s.\n"
            "Headers found in the sheet:\n  %s\n"
            "Check the column names in mapping.json -> source.headers."
            % (sheet, ", ".join('\"%s\"' % m for m in missing), "\n  ".join(found))
        )

    src_rows = all_rows[hdr_idx + 1:]
    by_sector = defaultdict(dict)   # (site_new, ne_id, Y) -> {X: e_tilt}
    sector_site_old = {}            # (site_new, ne_id, Y) -> site_old
    sector_order = []
    seen_sector = set()
    skipped = []

    def _get(r, field):
        j = col[field]
        return r[j] if (j is not None and len(r) > j) else None

    for r in src_rows:
        if not r or _get(r, "cell_name") is None:
            continue
        site_old = _get(r, "site_old")
        site_old = str(site_old).strip() if site_old is not None else ""
        site_new = _get(r, "site_new")
        ne_id = _get(r, "ne_id")
        cell = str(_get(r, "cell_name"))
        e_tilt = _get(r, "e_tilt")
        if len(cell) < 2:
            skipped.append((cell, "too short"))
            continue
        X = cell[-2]
        Y = cell[-1]
        if X not in band_map or X == "_NOT_USED_":
            skipped.append((cell, f"unknown X={X}"))
            continue
        if Y not in sector_map:
            skipped.append((cell, f"unknown Y={Y}"))
            continue
        key = (site_new, ne_id, Y)
        if key not in seen_sector:
            seen_sector.add(key)
            sector_order.append(key)
        sector_site_old[key] = site_old
        try:
            e_tilt_val = float(e_tilt) if e_tilt not in (None, "-") else 0.0
        except (TypeError, ValueError):
            e_tilt_val = 0.0
        by_sector[key][X] = e_tilt_val

    rows = []
    for key in sector_order:
        site_new, ne_id, Y = key
        sec = sector_map[Y]
        site_name = f"{site_new}_{ne_id}"
        sector_num = sec["sector_id"][1:]  # "S1" -> "1"
        site_old = sector_site_old.get(key)
        for X in DEVICE_ORDER:
            band = band_map[X]
            slot = band["slot_suffix"]
            band_token = band["band_token"]
            if X == "_NOT_USED_":
                rcu_tilt = 0
            elif X == tg_band:
                # Device No 3: presence + tilt come from the 3G sheet.
                if (site_old, sector_num) in three_g_tilts:
                    tg_tilt = three_g_tilts[(site_old, sector_num)]
                    rcu_tilt = round((tg_tilt or 0.0) * 10)
                else:
                    band_token = tg_absent
                    rcu_tilt = 0
            else:
                rcu_tilt = round(by_sector[key].get(X, 0.0) * 10)
            rru_name = f"{site_new}_{ne_id}_{band_token}_{sec['sector_id']}{slot}"
            rows.append([
                site_name,
                rru_name,
                rru_cn,
                sec["rru_srn"],
                rru_sn,
                band["color"],
                rcu_tilt,
                rru_name,   # column H = column B (Device Name = RRU Name)
            ])

    return rows, skipped, len(sector_order)


def _find_template_header_row(ws, headers, max_scan=20):
    """Find the column-header row in the template by matching known header tokens.

    Templates may have preamble rows above the header (e.g. an orange
    'Declaration' note), so the header is not always row 1. Returns the 1-based
    row index whose cells best match ``headers``; falls back to row 1.
    """
    targets = [_norm_header(h) for h in headers]
    max_col = ws.max_column or 1
    best_row, best_score = 1, 0
    for i in range(1, min(ws.max_row or 1, max_scan) + 1):
        cells = [_norm_header(ws.cell(row=i, column=c).value) for c in range(1, max_col + 1)]
        score = sum(1 for t in targets if t and any(t in cell for cell in cells))
        if score > best_score:
            best_score, best_row = score, i
    return best_row if best_score >= 2 else 1


def write_output(template_path, target_sheet, rows, output_path):
    """Write rows into a copy of the template, preserving its header rows and
    the styling of its first data row.

    The header row is detected (it may sit below preamble rows such as a
    Declaration note), so everything up to and including the header is kept and
    data is written starting on the next row.
    """
    out_wb = load_workbook(template_path)
    out_ws = out_wb[target_sheet]

    header_row = _find_template_header_row(out_ws, HEADERS)
    data_start = header_row + 1

    # Sample styling from the template's first data row (below the header);
    # if the template has no data rows, fall back to the header row.
    style_row = data_start if (out_ws.max_row or 0) >= data_start else header_row
    template_row_styles = []
    for col in range(1, out_ws.max_column + 1):
        cell = out_ws.cell(row=style_row, column=col)
        template_row_styles.append({
            "font": copy(cell.font),
            "fill": copy(cell.fill),
            "border": copy(cell.border),
            "alignment": copy(cell.alignment),
            "number_format": cell.number_format,
        })

    # Clear existing data rows only (keep preamble + header).
    if (out_ws.max_row or 0) >= data_start:
        out_ws.delete_rows(data_start, out_ws.max_row - data_start + 1)

    for i, values in enumerate(rows):
        out_row = data_start + i
        for col_idx, v in enumerate(values, start=1):
            c = out_ws.cell(row=out_row, column=col_idx, value=v)
            if col_idx - 1 < len(template_row_styles):
                s = template_row_styles[col_idx - 1]
                c.font = copy(s["font"])
                c.fill = copy(s["fill"])
                c.border = copy(s["border"])
                c.alignment = copy(s["alignment"])
                c.number_format = s["number_format"]

    out_wb.save(output_path)
