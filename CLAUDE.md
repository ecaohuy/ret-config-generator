# RET Config Generator — project rules

Tool that generates `RETConfigWDTInternal` output from a CDD workbook.
Entry points: `ret_gui.py` (Tkinter GUI), `generate_ret.py` (CLI); shared
logic in `ret_core.py`. Windows EXE is built by GitHub Actions on `v*` tags.

## Design rules (must follow)

1. **Any rule change goes in `mapping.json`** — transformation logic (column
   header names, sheet names, naming patterns, bands/colors, tilt sources)
   must be configurable in `mapping.json`. Do not hardcode such values in
   Python; if a new rule is requested, extend `mapping.json` (and the code
   that reads it) rather than embedding constants.
2. **Any style change goes in the template** — the output workbook's look
   (banner/preamble rows, header row, colors, fonts, formats) comes from the
   template file (`RETConfigWDTInternal.xlsx`). The generator must copy the
   template, preserve everything above and including the header row, and
   write only data rows, sampling their style from the template's first data
   row. Never style output cells from hardcoded values in code.

## Release flow

Commit → push → tag `v1.0.x` → CI builds `RET-Generator.exe` and attaches it
to the GitHub Release. `mapping.json` and the template are bundled into the
EXE via PyInstaller `--add-data`.
