"""Generate RETConfigWDTInternal_new.xlsx from CDD.xlsx using mapping.json.

Thin CLI wrapper around ret_core (the same logic the GUI uses), so column
resolution, the Device No 3 / 3G rule, and styling stay identical everywhere.
"""
import ret_core

CDD_PATH = "CDD.xlsx"
TEMPLATE_PATH = "RETConfigWDTInternal.xlsx"
MAPPING_PATH = "mapping.json"
OUTPUT_PATH = "RETConfigWDTInternal_new.xlsx"


def main():
    mapping = ret_core.load_mapping(MAPPING_PATH)
    sheet = mapping["source"]["sheet"]
    target_sheet = mapping.get("target", {}).get("sheet", "Internal")

    rows, skipped, sectors = ret_core.build_rows(CDD_PATH, sheet, mapping)
    ret_core.write_output(TEMPLATE_PATH, target_sheet, rows, OUTPUT_PATH)

    print(f"Wrote {OUTPUT_PATH}: {len(rows)} data rows across {sectors} sectors.")
    if skipped:
        print(f"Skipped {len(skipped)} CDD rows:")
        for cell, why in skipped[:20]:
            print(" -", cell, "->", why)
        if len(skipped) > 20:
            print(f"   ... +{len(skipped) - 20} more")


if __name__ == "__main__":
    main()
