# NewRET — project rules

Same idea as `../01.RET` but **much simpler**. Generates a
`RETConfigWDTInternal` workbook (Internal sheet) from a CDD workbook's
`4G CDD` sheet, plus the optional `ADD RET`/`MOD RETTILT` MML rewrite.

Entry points: `ret_gui.py` (Tkinter, two tabs), `generate_ret.py`
(CLI, CDD→WDT xlsx), `generate_ret_text.py` (CLI, RET_template.txt→
RET_output.txt MML). Shared logic in `ret_core.py`. Run with `uv run`.

## How it differs from ../01.RET (the simplifications)

- Only the **`4G CDD`** sheet is used. No 3G Installation Design sheet,
  no Local Cell ID band lookup.
- Each sector always emits **exactly 4 fixed RET devices** (mapping.json
  → `devices`): Lb1 `L1800 _1`, CLb2 `L1800 _2`, CRb3 `NSN_L1800 _1`,
  Rb4 `E///_U2100 _2`.
- Band (X) and sector (Y) come straight from `CellName (New)[Key]`:
  `X = LEFT(RIGHT(cell,2),1)`, `Y = RIGHT(cell,1)`. The band letter only
  picks which CDD row's `E_TILT` feeds each device's tilt (`tilt_band`,
  `tilt_fallback`): Lb1/CRb3/Rb4 use band C (1800 F1), CLb2 uses band D.
- RRU/Device Name = `{SiteName_New}_{band_token}_{sector_id}{slot}`
  (no Ne ID). Site Name(*) = `NEName_New` (verbatim from the CDD).
- Sectors are resolved by a computed rule (`sector_rule`), not a fixed
  table: digit `1-9` or letter `A-I` → base sector n; `rru_srn = srn_base +
  (n-1)` (S1=60…). A **co-located offset** then applies: if
  `LEFT(NEName_New, match_prefix_len)` == `LEFT(CellName, match_prefix_len)`
  (default `match_prefix_len`=8, the site code) the cell is the NE's own
  site and keeps n; otherwise it is a co-located neighbour and gets
  `+colocated_offset` (default 3: S1→S4, S2→S5, S3→S6). Non-conforming
  names (e.g. dashed `VNP-4G-…`) skip.
- New feature: **BBU Cluster filter**. The GUI multi-selects distinct
  `BBU Cluster` values; only those sectors are emitted (none = all).
- MML feature keeps the old behaviour: DEVICENAME prefix →
  `{SiteName_New}_{Ne ID}` (prefix_token_count=2), so the MML template's
  `{site}_{neid}_…` names are rewritten and tilts matched positionally.

## Design rules (inherited from ../01.RET — must follow)

1. **Any rule change goes in `mapping.json`** — header names, sheet name,
   band/sector rules, device list, colors, tilt sources, naming patterns.
   Do not hardcode such values in Python.
2. **Any style change goes in the template** (`Template.xlsx`). The
   generator copies the template, keeps everything up to and including the
   header row, and writes only data rows, sampling style from the
   template's first data row. Never style output cells from code.
