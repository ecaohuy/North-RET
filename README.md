# North-RET

Generates a `RETConfigWDTInternal` workbook from a CDD workbook's **`4G CDD`**
sheet, plus an optional `ADD RET` / `MOD RETTILT` MML rewrite. A simpler sibling
of the original RET Config Generator.

## What it does

For every sector in the `4G CDD` sheet it emits **4 fixed RET devices**:

| # | Color | RRU / Device Name              | Tilt source (E_TILT ×10) |
|---|-------|--------------------------------|--------------------------|
| 1 | Lb1   | `{Site}_L1800_S{n}_1`          | band C (1800 F1)         |
| 2 | CLb2  | `{Site}_L1800_S{n}_2`          | band D (1800 F2)         |
| 3 | CRb3  | `{Site}_NSN_L1800_S{n}_1`      | band C (1800 F1)         |
| 4 | Rb4   | `{Site}_E///_U2100_S{n}_2`     | band C (1800 F1)         |

- **Site** = `SiteName (RRU Location)_New`; Site Name(*) = `{Site}_LN`.
- Band/sector come from `CellName (New)[Key]`: band = 2nd-to-last char,
  sector = last char (digit `1-9` or letter `A-I` → sector n).
- **BBU Cluster filter**: pick one/more `BBU Cluster` values to limit the
  output (none = all).

## Run from source

```bash
uv run ret_gui.py            # GUI (two tabs)
uv run generate_ret.py       # CLI: CDD -> RETConfigWDTInternal_new.xlsx
uv run generate_ret.py "Cam Le 3"   # only that BBU Cluster
uv run generate_ret_text.py  # CLI: RET_template.txt -> RET_output.txt
```

## Windows EXE

Pushing a `v*` tag triggers GitHub Actions to build `North-RET.exe` and attach
it to the release.

## Design rules

1. **Transformation rules live in `mapping.json`** — header names, sheet,
   band/sector rules, device list, colors, tilt sources, naming patterns.
2. **Styling lives in the template** (`Template.xlsx`) — the generator copies it,
   keeps everything up to and including the header row, and writes only data
   rows, sampling style from the first data row.
