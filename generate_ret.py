"""Generate RETConfigWDTInternal_new.xlsx from CDD.xlsx (4G CDD) using mapping.json.

Thin CLI wrapper around ret_core (same logic the GUI uses). Optional BBU Cluster
filter: pass cluster names as extra args to limit the output to those clusters.

    uv run generate_ret.py                  # all clusters
    uv run generate_ret.py "Cam Le 3"       # only that BBU Cluster
"""
import sys

import ret_core

CDD_PATH = "CDD.xlsx"
TEMPLATE_PATH = "Template.xlsx"
MAPPING_PATH = "mapping.json"
OUTPUT_PATH = "RETConfigWDTInternal_new.xlsx"


def main():
    mapping = ret_core.load_mapping(MAPPING_PATH)
    sheet = mapping["source"]["sheet"]
    target_sheet = mapping.get("target", {}).get("sheet", "Internal")
    clusters = sys.argv[1:] or None

    rows, skipped, sectors, _ = ret_core.build_rows(CDD_PATH, sheet, mapping, clusters=clusters)
    ret_core.write_output(TEMPLATE_PATH, target_sheet, rows, OUTPUT_PATH)

    scope = f" (clusters: {', '.join(clusters)})" if clusters else ""
    print(f"Wrote {OUTPUT_PATH}: {len(rows)} data rows across {sectors} sectors{scope}.")
    if skipped:
        print(f"Skipped {len(skipped)} CDD rows:")
        for cell, why in skipped[:20]:
            print(" -", cell, "->", why)
        if len(skipped) > 20:
            print(f"   ... +{len(skipped) - 20} more")


if __name__ == "__main__":
    main()
