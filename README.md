# RET Config Generator

A small desktop tool (Tkinter GUI) that generates `RETConfigWDTInternal` from a
CDD workbook using the rules in `mapping.json`.

## Download (Windows)

Grab the latest `RET-Generator.exe` from the **[Releases page](../../releases/latest)**.
No install needed — double-click to run.

## How it works

1. **Browse** to your `CDD.xlsx` and pick the `4G CDD` sheet.
2. **Validate / Preview** to check the generated rows.
3. **Generate** to write the output workbook.

`mapping.json` and the style template are bundled into the `.exe`, so the tool
works out of the box; you only need to provide your own CDD file at runtime.

## Design rules

- **Logic lives in `mapping.json`** — any change to the transformation rules
  (column names, naming patterns, bands, tilt sources, …) is made by editing
  `mapping.json`, never by hardcoding values in the Python code.
- **Style lives in the template** — any change to the output's look (header
  rows, colors, fonts, banner text) is made by editing the template workbook
  (`RETConfigWDTInternal.xlsx`). The generator copies the template, keeps its
  preamble/header rows, and only fills in data rows, sampling their style from
  the template's first data row.

### Device No 3 rule
Device No 3 (band `E` / `NSN_U2100`) is resolved against the **3G Installation
Design** sheet. If the sector's Logical Sector Name (`{SiteName_Old}-{sector#}`)
exists there, the row is `…_NSN_U2100_S{n}_1` with tilt from that sheet's
*New E-Tilt*; otherwise it becomes `…_NOT_USED_S{n}_1` with tilt 0.

## Build locally

```bash
uv run ret_gui.py      # GUI
uv run generate_ret.py # CLI
```

The Windows `.exe` is built automatically by GitHub Actions
(`.github/workflows/build-exe.yml`) on every `v*` tag.
