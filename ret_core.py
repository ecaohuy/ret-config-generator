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
    sec_col = ord(cfg.get("logical_sector_column", "B").upper()) - ord("A")
    tilt_col = ord(cfg.get("etilt_column", "F").upper()) - ord("A")
    sep = cfg.get("separator", "-")
    tilts = {}
    for r in ws.iter_rows(min_row=2, values_only=True):
        if len(r) <= sec_col or r[sec_col] is None:
            continue
        site_old, _, num = str(r[sec_col]).partition(sep)
        if not num:
            continue
        key = (site_old.strip(), num.strip())
        e_tilt = r[tilt_col] if len(r) > tilt_col else None
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
    src_rows = list(cdd_ws.iter_rows(min_row=2, values_only=True))
    three_g_tilts = load_three_g_tilts(cdd_wb, mapping)
    cdd_wb.close()

    by_sector = defaultdict(dict)   # (site_new, ne_id, Y) -> {X: e_tilt}
    sector_site_old = {}            # (site_new, ne_id, Y) -> site_old
    sector_order = []
    seen_sector = set()
    skipped = []

    for r in src_rows:
        if not r or len(r) < 7 or r[6] is None:
            continue
        site_old = r[0]
        site_new = r[1]
        ne_id = r[5]
        cell = str(r[6])
        e_tilt = r[7] if len(r) > 7 else None
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


def write_output(template_path, target_sheet, rows, output_path):
    """Write rows into a copy of the template, preserving row-2 styling."""
    out_wb = load_workbook(template_path)
    out_ws = out_wb[target_sheet]

    template_row_styles = []
    for col in range(1, out_ws.max_column + 1):
        cell = out_ws.cell(row=2, column=col)
        template_row_styles.append({
            "font": copy(cell.font),
            "fill": copy(cell.fill),
            "border": copy(cell.border),
            "alignment": copy(cell.alignment),
            "number_format": cell.number_format,
        })

    if out_ws.max_row >= 2:
        out_ws.delete_rows(2, out_ws.max_row - 1)

    for i, values in enumerate(rows):
        out_row = 2 + i
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
