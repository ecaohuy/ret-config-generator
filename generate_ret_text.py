"""Generate RET_output.txt from RET_template.txt + RET_input.txt + CDD.xlsx.

Thin CLI wrapper around ret_core.write_text_output so the DEVICENAME prefix,
SERIALNO and TILT rewrite rules stay identical in the GUI and CLI.
"""
import ret_core

INPUT_PATH = "RET_input.txt"
TEMPLATE_PATH = "RET_template.txt"
CDD_PATH = "CDD.xlsx"
MAPPING_PATH = "mapping.json"
OUTPUT_PATH = "RET_output.txt"


def main():
    mapping = ret_core.load_mapping(MAPPING_PATH)
    sheet = mapping["source"]["sheet"]

    warnings = ret_core.write_text_output(
        TEMPLATE_PATH, INPUT_PATH, CDD_PATH, sheet, mapping, OUTPUT_PATH
    )

    print(f"Wrote {OUTPUT_PATH}.")
    if warnings:
        print(f"{len(warnings)} warning(s):")
        for w in warnings[:20]:
            print(" -", w)
        if len(warnings) > 20:
            print(f"   ... +{len(warnings) - 20} more")


if __name__ == "__main__":
    main()
