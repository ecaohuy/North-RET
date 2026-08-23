"""NewRET core: CDD '4G CDD' sheet -> RETConfigWDTInternal (Internal sheet).

Shared by generate_ret.py (CLI), generate_ret_text.py (CLI) and ret_gui.py.

Simpler than ../01.RET: no 3G Installation Design sheet, no Local Cell ID band
lookup, and a FIXED set of 4 RET devices per sector (mapping.json -> devices).
Band (X) and sector (Y) come straight from CellName (New)[Key]:
    X = LEFT(RIGHT(CellName, 2), 1)   (band letter)
    Y = RIGHT(CellName, 1)            (sector letter)
All transformation rules live in mapping.json; styling comes from the template.
"""
import json
import os
import re
from collections import defaultdict
from copy import copy

from openpyxl import load_workbook

import fast_xlsx

# Output column headers (target order), same 8 columns as the WDT Internal sheet.
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


def mapping_problems(mapping):
    """Return a list of human-readable reasons ``mapping`` is not NewRET format.

    Empty list = a valid NewRET mapping. The GUI lets the user Browse to any
    mapping.json; pointing it at the older ../01.RET mapping (which has no
    ``devices``/``constants``) used to crash with a bare ``KeyError: 'devices'``.
    """
    problems = []
    if not isinstance(mapping.get("devices"), list) or not mapping.get("devices"):
        problems.append('"devices" (the list of 4 RET devices per sector)')
    consts = mapping.get("constants")
    if not isinstance(consts, dict) or "rru_cn" not in consts or "rru_sn" not in consts:
        problems.append('"constants" with "rru_cn" and "rru_sn"')
    return problems


def _require_mapping_keys(mapping):
    """Raise a clear ValueError if ``mapping`` is not a NewRET mapping."""
    problems = mapping_problems(mapping)
    if problems:
        raise ValueError(
            "This mapping.json is not a NewRET mapping — it is missing: "
            + "; ".join(problems)
            + ".\nSelect the NewRET mapping.json (the one bundled with this app), "
            "not the mapping.json from the older RET project."
        )


def _norm_header(s):
    """Normalize a header cell for tolerant matching (case/whitespace-insensitive)."""
    if s is None:
        return ""
    return " ".join(str(s).split()).strip().lower()


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

    The 4G CDD sheet has a category row above the real header row, so the header
    is not row 0. Scans the first ``max_scan`` rows; falls back to row 0.
    """
    for i, row in enumerate(all_rows[:max_scan]):
        norm = [_norm_header(h) for h in row]
        if _find_index(norm, key_targets) is not None:
            return i, norm
    return 0, [_norm_header(h) for h in (all_rows[0] if all_rows else ())]


def _to_float(v):
    """Best-effort float; None for blank/'-' or unparseable."""
    if v in (None, "", "-"):
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def resolve_sector(Y, mapping, offset=0):
    """Map the sector letter/digit Y (+ offset) to (sector_id, rru_srn), or None.

    Per mapping.json -> sector_rule: a digit 1-9 is that base sector number; a
    single letter A-I maps to 1-9 (A=S1 ... I=S9). ``offset`` (the co-located
    offset, 0 for the NE's own site) is added to that base number. sector_id =
    'S{n}', rru_srn = srn_base + (n - 1). Returns None for anything outside
    1..max_sectors.
    """
    rule = mapping.get("sector_rule", {})
    base = rule.get("srn_base", 60)
    max_sectors = rule.get("max_sectors", 9)
    Y = str(Y).strip()
    n = None
    if Y.isdigit():
        n = int(Y)
    elif len(Y) == 1 and Y.upper().isalpha():
        n = ord(Y.upper()) - ord("A") + 1
    if n is None:
        return None
    n += offset
    if not (1 <= n <= max_sectors):
        return None
    return f"S{n}", base + (n - 1)


def resolve_row_sector(Y, lsid, offset, mapping):
    """Resolve a CDD row to (sector_id, rru_srn), honouring the digit exception.

    Per mapping.json -> sector_rule.digit_from_logical_sector_id: when
    RIGHT(CellName, 1) (=Y) is a DIGIT, the letter/co-located-offset rule does
    not apply; the sector is taken from RIGHT(Logical Sector ID (Site), 1)
    instead (e.g. LSID '3.1' -> S1, '3.2' -> S2, '3' -> S3), with no offset.
    Falls back to the plain digit-of-CellName rule when the LSID is unusable.
    Otherwise (letter Y) defers to ``resolve_sector`` with ``offset``.
    """
    rule = mapping.get("sector_rule", {})
    Y = str(Y).strip()
    if rule.get("digit_from_logical_sector_id", True) and Y.isdigit():
        if lsid not in (None, ""):
            last = str(lsid).strip()[-1:]
            if last.isdigit():
                return resolve_sector(last, mapping, 0)
    return resolve_sector(Y, mapping, offset)


def list_sheets(cdd_path):
    """Return sheet names in the CDD workbook."""
    try:
        with fast_xlsx.Workbook(cdd_path) as wb:
            return list(wb.sheets)
    except Exception:
        wb = load_workbook(cdd_path, read_only=True)
        try:
            return wb.sheetnames
        finally:
            wb.close()


def _wanted_headers(mapping):
    """The CDD header name expected for each field, from mapping.json."""
    hdr_cfg = mapping.get("source", {}).get("headers", {})
    return {
        "site_new": hdr_cfg.get("site_new", "SiteName (RRU Location)_New"),
        "ne_name": hdr_cfg.get("ne_name", "NEName_New"),
        "ne_id": hdr_cfg.get("ne_id", "Ne ID (New)"),
        "cell_name": hdr_cfg.get("cell_name", "CellName (New)[Key]"),
        "e_tilt": hdr_cfg.get("e_tilt", "E_TILT"),
        "bbu_cluster": hdr_cfg.get("bbu_cluster", "BBU Cluster"),
        "site_type": hdr_cfg.get("site_type", "Site Type"),
        "logical_sector_id": hdr_cfg.get("logical_sector_id", "Logical Sector ID (Site)"),
    }


def _resolve_columns(all_rows, mapping):
    """Locate CDD columns by header name. Returns (header_row_index, col_map, wanted)."""
    wanted = _wanted_headers(mapping)
    hdr_idx, norm = _locate_header(all_rows, [_norm_header(wanted["cell_name"])])
    col = {f: _find_index(norm, [_norm_header(name)]) for f, name in wanted.items()}
    return hdr_idx, col, wanted


# --------------------------------------------------------------------------
# CDD reading.
#
# The production CDD is ~22 MB / 15 700 rows x 106 columns and openpyxl needs
# ~9 s to materialise it, which used to be paid on every Preview/Generate.
# ``read_cdd`` instead pulls only the ~8 mapped columns straight out of the
# sheet XML (fast_xlsx, ~0.5 s) and memoises the result per (file, mtime,
# size, sheet, header names), so repeat runs are instant. Any surprise in the
# workbook falls back to the original openpyxl path.
# --------------------------------------------------------------------------

_CDD_CACHE = {}          # cache key -> (hdr_idx, col, wanted, records)
_CDD_CACHE_MAX = 3       # keep a couple of workbooks, not every one ever opened


class CDDData:
    """Parsed CDD sheet: mapped columns only, one dict of values per data row."""

    __slots__ = ("hdr_idx", "col", "wanted", "records", "header_names")

    def __init__(self, hdr_idx, col, wanted, records, header_names):
        self.hdr_idx = hdr_idx
        self.col = col
        self.wanted = wanted
        self.records = records
        self.header_names = header_names

    def missing(self, fields):
        """Header names of ``fields`` that were not found in the sheet."""
        return [self.wanted[f] for f in fields if self.col.get(f) is None]


def _records_from_rows(all_rows, hdr_idx, col):
    """Generic reader path: full row tuples -> per-row {field: value} dicts."""
    fields = [(f, j) for f, j in col.items() if j is not None]
    records = []
    for r in all_rows[hdr_idx + 1:]:
        if not r:
            continue
        rec = {}
        for f, j in fields:
            v = r[j] if len(r) > j else None
            if v is not None and v != "":
                rec[f] = v
        if rec:
            records.append(rec)
    return records


def _read_cdd_fast(cdd_path, sheet, mapping):
    with fast_xlsx.Workbook(cdd_path) as wb:
        if sheet not in wb.sheets:
            raise KeyError(sheet)
        head = wb.head_rows(sheet, 10)
        if not any(head):
            # Nothing decoded at all -> the sheet XML is not shaped the way this
            # reader expects; let the caller fall back to openpyxl.
            raise ValueError("fast reader decoded no header cells")
        hdr_idx, col, wanted = _resolve_columns(head, mapping)
        header_names = [str(h) for h in head[hdr_idx] if h is not None]
        if col.get("cell_name") is None:
            # Header genuinely absent (e.g. the 5G sheet). Return the empty
            # column map so build_rows can report it immediately.
            return CDDData(hdr_idx, col, wanted, [], header_names)
        indices = [j for j in col.values() if j is not None]
        raw = wb.read_columns(sheet, indices, skip_rows=hdr_idx + 1)
    by_index = {j: f for f, j in col.items() if j is not None}
    records = [{by_index[j]: v for j, v in row.items() if j in by_index} for row in raw]
    return CDDData(hdr_idx, col, wanted, [r for r in records if r], header_names)


def _read_cdd_openpyxl(cdd_path, sheet, mapping):
    wb = load_workbook(cdd_path, data_only=True, read_only=True)
    try:
        ws = wb[sheet]
        all_rows = list(ws.iter_rows(min_row=1, values_only=True))
    finally:
        wb.close()
    hdr_idx, col, wanted = _resolve_columns(all_rows, mapping)
    header_names = [str(h) for h in (all_rows[hdr_idx] if all_rows else ()) if h is not None]
    return CDDData(hdr_idx, col, wanted, _records_from_rows(all_rows, hdr_idx, col),
                   header_names)


def read_cdd(cdd_path, sheet, mapping):
    """Return the ``CDDData`` for ``sheet``, memoised on the file's mtime/size."""
    wanted = _wanted_headers(mapping)
    try:
        st = os.stat(cdd_path)
        key = (os.path.abspath(cdd_path), st.st_mtime_ns, st.st_size, sheet,
               tuple(sorted(wanted.items())))
    except OSError:
        key = None
    if key is not None and key in _CDD_CACHE:
        return _CDD_CACHE[key]
    try:
        data = _read_cdd_fast(cdd_path, sheet, mapping)
    except Exception:
        data = _read_cdd_openpyxl(cdd_path, sheet, mapping)
    if key is not None:
        if len(_CDD_CACHE) >= _CDD_CACHE_MAX:
            _CDD_CACHE.pop(next(iter(_CDD_CACHE)))
        _CDD_CACHE[key] = data
    return data


def clear_cdd_cache():
    """Drop memoised CDD parses (used by the GUI's 'Reload CDD')."""
    _CDD_CACHE.clear()


def list_bbu_clusters(cdd_path, sheet, mapping):
    """Return the sorted distinct BBU Cluster values present in the sheet."""
    data = read_cdd(cdd_path, sheet, mapping)
    if data.col.get("bbu_cluster") is None:
        return []
    found = set()
    for rec in data.records:
        if "cell_name" not in rec:
            continue
        v = rec.get("bbu_cluster")
        if v not in (None, ""):
            found.add(str(v).strip())
    return sorted(found)


def build_rows(cdd_path, sheet, mapping, clusters=None):
    """Parse the 4G CDD sheet and produce output rows.

    ``clusters`` (optional): an iterable of BBU Cluster values; when given, only
    sectors whose BBU Cluster is in the set are emitted. None/empty = all.

    Returns (rows, skipped, sector_count, site_index) where:
      rows         = list of 8-value lists (matching HEADERS)
      skipped      = list of (cell_name, reason)
      sector_count = number of distinct (site, sector) groups emitted
      site_index   = {site_new: {"prefix": ..., "tilts": {(sector_id, pos): tilt}}}
                     for the MML text feature.
    """
    _require_mapping_keys(mapping)
    devices = mapping["devices"]
    rru_cn = mapping["constants"]["rru_cn"]
    rru_sn = mapping["constants"]["rru_sn"]
    field_rules = mapping.get("field_rules", {})
    # Site Name(*) comes from this CDD field (default NEName_New).
    site_name_source = field_rules.get("site_name", {}).get("source", "ne_name")
    # RRU Name(*) prefix = the part of this field before the first of these delimiters.
    rru_prefix_source = field_rules.get("rru_name", {}).get("prefix_source", "site_new")
    rru_prefix_delims = field_rules.get("rru_name", {}).get("prefix_delimiters", ["_"])
    sector_rule = mapping.get("sector_rule", {})
    match_len = sector_rule.get("match_prefix_len", 8)
    colocated_offset = sector_rule.get("colocated_offset", 3)
    cluster_filter = {str(c).strip() for c in clusters} if clusters else None
    # Row filter: drop rows whose Site Type is one of skip_values (case-insensitive,
    # exact). Default skips a standalone "IBC" (but not "Macro+IBC", "CRAN-IBC", ...).
    site_type_cfg = mapping.get("row_filters", {}).get("site_type", {})
    site_type_skip = {str(v).strip().lower() for v in site_type_cfg.get("skip_values", [])}
    # RET MML only: which CDD field the RET input line 1 is matched against AND
    # which supplies the SITE token of the rewritten DEVICENAME prefix. "ne_name"
    # = NEName_New (default), "site_new" = SiteName_New.
    site_match = mapping.get("text_config", {}).get("site_match", {})
    site_match_field = site_match.get("field", "ne_name")
    # Append the Ne ID to the prefix? Templates whose DEVICENAME is just
    # {site}_{band}_{sector}_{slot} (no Ne ID) set this false.
    include_ne_id = site_match.get("include_ne_id", False)

    data = read_cdd(cdd_path, sheet, mapping)
    missing = data.missing(("site_new", "cell_name"))
    if missing:
        raise ValueError(
            "Sheet '%s': could not find required column(s) by header name: %s.\n"
            "Headers found in the sheet:\n  %s\n"
            "Check mapping.json -> source.headers."
            % (sheet, ", ".join('"%s"' % m for m in missing), "\n  ".join(data.header_names))
        )

    def _get(r, field):
        return r.get(field)

    by_sector = defaultdict(dict)   # (site_new, Y) -> {X: e_tilt}
    sector_ne_id = {}               # (site_new, Y) -> Ne ID (first non-blank)
    sector_ne_name = {}             # (site_new, Y) -> NEName_New (first non-blank)
    sector_offset = {}              # (site_new, Y) -> 0 (own site) or colocated_offset
    sector_lsid = {}                # (site_new, Y) -> Logical Sector ID (first non-blank)
    sector_order = []
    seen = set()
    skipped = []

    for r in data.records:
        if _get(r, "cell_name") is None:
            continue
        cell = str(_get(r, "cell_name")).strip()
        if len(cell) < 2:
            skipped.append((cell, "too short"))
            continue
        site_new = _get(r, "site_new")
        site_new = str(site_new).strip() if site_new is not None else ""
        if not site_new:
            skipped.append((cell, "blank SiteName_New"))
            continue
        # Site Type filter: e.g. a standalone "IBC" is ignored entirely.
        if site_type_skip:
            st = _get(r, "site_type")
            st = str(st).strip().lower() if st is not None else ""
            if st in site_type_skip:
                skipped.append((cell, f"Site Type={st or '(blank)'} (filtered)"))
                continue
        # BBU Cluster filter (when a selection is active).
        if cluster_filter is not None:
            bbu = _get(r, "bbu_cluster")
            bbu = str(bbu).strip() if bbu is not None else ""
            if bbu not in cluster_filter:
                continue
        X = cell[-2]   # band letter
        Y = cell[-1]   # sector letter/digit
        # Own site (cell belongs to this NE) keeps its base sector number; a
        # co-located neighbour (LEFT(NEName,N) != LEFT(CellName,N)) is shifted by
        # colocated_offset (S1->S4, S2->S5, S3->S6).
        ne_name_val = _get(r, "ne_name")
        ne_name_val = str(ne_name_val).strip() if ne_name_val is not None else ""
        is_own = ne_name_val[:match_len].upper() == cell[:match_len].upper()
        offset = 0 if is_own else colocated_offset
        # When RIGHT(CellName,1) is a digit, the sector comes from the Logical
        # Sector ID (Site) instead of the letter/offset rule (see resolve_row_sector).
        lsid_val = _get(r, "logical_sector_id")
        if resolve_row_sector(Y, lsid_val, offset, mapping) is None:
            skipped.append((cell, f"unknown sector Y={Y}"))
            continue
        key = (site_new, Y)
        if key not in seen:
            seen.add(key)
            sector_order.append(key)
            sector_offset[key] = offset
        if lsid_val not in (None, "") and key not in sector_lsid:
            sector_lsid[key] = lsid_val
        ne_id = _get(r, "ne_id")
        if ne_id not in (None, "") and key not in sector_ne_id:
            sector_ne_id[key] = ne_id
        nn = _get(r, site_name_source)
        if nn not in (None, "") and key not in sector_ne_name:
            sector_ne_name[key] = str(nn).strip()
        # E_TILT per band; first non-blank wins so a blank row can't clobber it.
        t = _to_float(_get(r, "e_tilt"))
        if X not in by_sector[key] or (by_sector[key].get(X) is None and t is not None):
            by_sector[key][X] = t

    rows = []
    site_index = {}
    for key in sector_order:
        site_new, Y = key
        sector_id, rru_srn = resolve_row_sector(
            Y, sector_lsid.get(key), sector_offset.get(key, 0), mapping)
        present = by_sector[key]            # {X: e_tilt} for this sector
        ne_id = sector_ne_id.get(key, "")
        # Site Name(*) = NEName_New (fallback to SiteName_New if blank).
        site_name = sector_ne_name.get(key, "") or site_new
        # RRU Name(*) prefix = part of the configured source field before the first
        # delimiter char (default SiteName (RRU Location)_New).
        rru_prefix = {"site_new": site_new, "ne_name": site_name,
                      "site_name": site_name}.get(rru_prefix_source, site_new)
        for d in rru_prefix_delims:
            rru_prefix = rru_prefix.split(d)[0]
        # RET MML: the RET input line 1 is matched against site_match_field
        # (NEName_New by default), which also supplies the DEVICENAME site token.
        prefix_site = site_name if site_match_field == "ne_name" else site_new
        prefix = f"{prefix_site}_{ne_id}".rstrip("_") if include_ne_id else str(prefix_site)
        site_entry = site_index.setdefault(
            prefix_site, {"prefix": prefix, "tilts": {}, "conflicts": {}}
        )
        # Two CellName groups of one site can resolve to the SAME sector (e.g.
        # a site whose Logical Sector IDs are 1.1/1.2/2.1/2.2 -> S1,S2,S1,S2).
        # Record it instead of letting the later group silently overwrite the
        # earlier one's tilts; the MML step surfaces it as a warning.
        claimed = site_entry.setdefault("sectors", {})
        prior_y = claimed.get(sector_id)
        if prior_y is None:
            claimed[sector_id] = Y
        elif prior_y != Y:
            site_entry["conflicts"].setdefault(sector_id, [prior_y]).append(Y)

        for pos, dev in enumerate(devices):
            t = present.get(dev["tilt_band"])
            if t is None:
                t = present.get(dev.get("tilt_fallback"))
            rcu_tilt = round((t or 0.0) * 10)
            rru_name = f"{rru_prefix}_{dev['band_token']}_{sector_id}{dev['slot_suffix']}"
            # First CellName group to claim a sector wins, so a later colliding
            # group cannot silently replace its tilts (see "conflicts" above).
            site_entry["tilts"].setdefault((sector_id, pos), rcu_tilt)
            rows.append([
                site_name,
                rru_name,
                rru_cn,
                rru_srn,
                rru_sn,
                dev["color"],
                rcu_tilt,
                rru_name,   # column H = column B (Device Name = RRU Name)
            ])

    return rows, skipped, len(sector_order), site_index


def _find_template_header_row(ws, headers, max_scan=20):
    """Find the column-header row in the template by matching known header tokens.

    Templates may have preamble rows above the header (e.g. a Declaration note),
    so the header is not always row 1. Returns the 1-based row index; falls back
    to row 1.
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
    """Write rows into a copy of the template, preserving its header rows and the
    styling of its first data row.

    The header row is detected (it may sit below preamble rows), so everything up
    to and including the header is kept and data is written from the next row.
    """
    out_wb = load_workbook(template_path)
    out_ws = out_wb[target_sheet]

    header_row = _find_template_header_row(out_ws, HEADERS)
    data_start = header_row + 1

    style_row = data_start if (out_ws.max_row or 0) >= data_start else header_row
    template_row_styles = []
    for c in range(1, out_ws.max_column + 1):
        cell = out_ws.cell(row=style_row, column=c)
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


# --------------------------------------------------------------------------
# RET_template.txt -> RET_output.txt (MML script) conversion.
#
# Rewrites three things in an "ADD RET / MOD RETTILT" MML template using the
# RET_input.txt serials and the CDD-derived rows (build_rows):
#   * DEVICENAME : replace the leading site token(s) with the site_match prefix
#                  (NEName_New by default, no Ne ID), keeping the band/sector/slot
#                  suffix so the output matches the template. The RET input line 1
#                  is matched against that same field (site_match.field).
#   * SERIALNO   : take the input serial matched by (CTRLSRN, RIGHT(serial, N)).
#   * TILT       : RCU Tilt of the same-DEVICENO ADD RET device, matched
#                  positionally (CTRLSRN -> sector, device order within sector).
# All rules/field names live in mapping.json -> text_config.
# --------------------------------------------------------------------------

_RE_DEVICENO = re.compile(r"DEVICENO=\s*(\d+)")
_RE_DEVICENAME = re.compile(r'DEVICENAME="([^"]*)"')
_RE_CTRLSRN = re.compile(r"CTRLSRN=\s*(\d+)")
_RE_SERIALNO = re.compile(r'SERIALNO="([^"]*)"')
_RE_TILT = re.compile(r"TILT=\s*(-?\d+)")


def parse_ret_input(input_path, mapping):
    """Parse RET_input.txt -> (site, serials)."""
    with open(input_path, encoding="utf-8") as f:
        text = f.read()
    return parse_ret_input_text(text, mapping, source=input_path)


def parse_ret_input_text(text, mapping, source="pasted input"):
    """Parse in-memory RET input text -> (site, serials).

    Line 1 (first non-blank) is the site name. Each remaining line is split on
    whitespace; ``serials`` maps (CTRLSRN, RIGHT(serial, suffix_len)) -> full
    serial, the key used to rewrite each template SERIALNO.
    """
    cfg = mapping.get("text_config", {}).get("input", {})
    srn_col = cfg.get("ctrlsrn_column", 1)
    ser_col = cfg.get("serial_column", 7)
    suf_len = cfg.get("serial_suffix_len", 3)

    site = None
    serials = {}
    for ln in text.splitlines():
        if not ln.strip():
            continue
        if site is None:
            site = ln.strip()
            continue
        parts = ln.split()
        if len(parts) <= max(srn_col, ser_col):
            continue
        srn = parts[srn_col].strip()
        serial = parts[ser_col].strip()
        serials[(srn, serial[-suf_len:])] = serial
    if site is None:
        raise ValueError("RET input is empty: %s" % source)
    return site, serials


def _site_entry(site_index, site):
    """Return the build_rows site entry for ``site``, matched tolerantly.

    The RET input's line 1 is typed/pasted by hand, so match case-insensitively
    and — failing that — on the bare site code (the part before the first '_'
    or '-'), which is what distinguishes 'HNIVTH16', 'HNIVTH16_LN' and
    'hnivth16_ln'. An ambiguous short form is reported rather than guessed.
    """
    entry = site_index.get(site)
    if entry is not None:
        return entry

    def _code(s):
        return re.split(r"[_-]", str(s).strip(), 1)[0].upper()

    want = str(site).strip().upper()
    ci = [k for k in site_index if k.upper() == want]
    if len(ci) == 1:
        return site_index[ci[0]]
    by_code = [k for k in site_index if _code(k) == _code(site)]
    if len(by_code) == 1:
        return site_index[by_code[0]]
    if len(by_code) > 1:
        raise ValueError(
            "Site %r (from the RET input) matches %d CDD sites: %s.\n"
            "Use the full NEName_New so the right one can be picked."
            % (site, len(by_code), ", ".join(sorted(by_code)))
        )
    near = sorted(k for k in site_index if _code(k).startswith(_code(site)[:6]))
    hint = ("\nDid you mean: %s" % ", ".join(near[:10])) if near else (
        "\n%d sites are available in this CDD/cluster selection." % len(site_index))
    raise ValueError(
        "Site %r (from the RET input) was not found in the CDD output.%s" % (site, hint)
    )


def _site_tilt_index(site_index, site):
    """Return (new_prefix, tilt_by_sector_pos) for ``site`` from build_rows."""
    entry = _site_entry(site_index, site)
    return entry["prefix"], entry["tilts"]


def _device_signatures(mapping):
    """[(band_token, slot_suffix, position)] longest band_token first.

    Used to read a template DEVICENAME suffix ('NSN_L1800_S1_1') back into the
    device it names, so a tilt is matched by *identity* rather than by the order
    the ADD RET lines happen to appear in.
    """
    sigs = []
    for pos, dev in enumerate(mapping.get("devices", [])):
        sigs.append((str(dev.get("band_token", "")), str(dev.get("slot_suffix", "")), pos))
    sigs.sort(key=lambda s: len(s[0]), reverse=True)
    return sigs


def _parse_devicename(suffix, sigs):
    """DEVICENAME suffix -> (sector_id, device position), or (None, None).

    ``suffix`` is the DEVICENAME with the leading site token(s) removed, e.g.
    'NSN_L1800_S1_1' -> ('S1', 2).
    """
    for band, slot, pos in sigs:
        head, tail = band + "_", slot
        if band and suffix.startswith(head) and (not tail or suffix.endswith(tail)):
            sector = suffix[len(head):len(suffix) - len(tail)]
            if sector:
                return sector, pos
    return None, None


def _tilt_line_prefixes(tcfg):
    """The MML verbs whose TILT= is rewritten.

    Accepts ``tilt_line_prefix`` as a string or a list; a Huawei template may
    carry the tilt on ``MOD RETTILT`` or on ``MOD RETSUBUNIT`` (or both), and
    only rewriting one of them silently leaves the other at the template's
    values, which is the worst possible failure — a plausible-looking script
    with somebody else's tilts.
    """
    raw = tcfg.get("tilt_line_prefix", tcfg.get("tilt_line_prefixes",
                                                ["MOD RETTILT", "MOD RETSUBUNIT"]))
    if isinstance(raw, str):
        raw = [raw]
    return tuple(str(p) for p in raw if str(p).strip())


def build_text_output(template_path, input_path, cdd_path, sheet, mapping,
                      input_text=None, clusters=None):
    """Produce the rewritten MML text. Returns (text, warnings, report)."""
    if input_text is not None and input_text.strip():
        site, serials = parse_ret_input_text(input_text, mapping)
    else:
        site, serials = parse_ret_input(input_path, mapping)
    _rows, _, _, site_index = build_rows(cdd_path, sheet, mapping, clusters=clusters)
    entry = _site_entry(site_index, site)
    new_prefix = entry["prefix"]
    tilt_by_sector_pos = entry["tilts"]

    # Map CTRLSRN -> sector_id by inverting the sector_rule over 1..max_sectors.
    max_sectors = mapping.get("sector_rule", {}).get("max_sectors", 9)
    srn_to_sector = {}
    for n in range(1, max_sectors + 1):
        sid, srn = resolve_sector(str(n), mapping)
        srn_to_sector[str(srn)] = sid

    tcfg = mapping.get("text_config", {}).get("template", {})
    prefix_tokens = tcfg.get("prefix_token_count", 2)
    add_prefix = tcfg.get("add_line_prefix", "ADD RET")
    tilt_prefixes = _tilt_line_prefixes(tcfg)
    suf_len = mapping.get("text_config", {}).get("input", {}).get("serial_suffix_len", 3)
    sigs = _device_signatures(mapping)

    with open(template_path, encoding="utf-8") as f:
        tmpl_lines = f.readlines()

    warnings = []
    report = []          # one dict per rewritten ADD RET line, for the preview
    used_serials = set()

    # Sectors the CDD gave this site twice (e.g. Logical Sector IDs 1.1/1.2 and
    # 2.1/2.2 both landing on S1/S2). The first group's tilts are used; say so.
    for sid, ys in sorted(entry.get("conflicts", {}).items()):
        warnings.append(
            "Sector %s is claimed by %d CellName groups of %s (last chars %s) — "
            "the first one's tilt is used." % (sid, len(ys), site, "/".join(map(str, ys)))
        )

    # Pass 1: DEVICENO -> RCU Tilt. The device is identified from its DEVICENAME
    # ({band_token}_{sector}{slot_suffix}), so the tilt follows the device even
    # if the ADD RET lines are reordered or a sector is missing devices; the old
    # order-of-appearance rule is only the fallback.
    deviceno_tilt = {}
    add_pos = defaultdict(int)
    for ln in tmpl_lines:
        if not ln.lstrip().startswith(add_prefix):
            continue
        m_no, m_srn = _RE_DEVICENO.search(ln), _RE_CTRLSRN.search(ln)
        if not m_no:
            continue
        srn = m_srn.group(1) if m_srn else None
        pos_seen = add_pos[srn]
        add_pos[srn] += 1
        srn_sector = srn_to_sector.get(srn) if srn is not None else None

        sector_id = pos = None
        m_dn = _RE_DEVICENAME.search(ln)
        if m_dn:
            suffix = "_".join(m_dn.group(1).split("_")[prefix_tokens:])
            sector_id, pos = _parse_devicename(suffix, sigs)
        matched_by = "DEVICENAME"
        if sector_id is None:
            sector_id, pos, matched_by = srn_sector, pos_seen, "CTRLSRN+order"
        elif srn_sector is not None and sector_id != srn_sector:
            warnings.append(
                "DEVICENO=%s: DEVICENAME says %s but CTRLSRN=%s means %s — "
                "using %s." % (m_no.group(1), sector_id, srn, srn_sector, sector_id)
            )

        tilt = tilt_by_sector_pos.get((sector_id, pos))
        if tilt is None:
            warnings.append(
                "No RCU Tilt in the CDD for DEVICENO=%s (sector=%s, device #%s, "
                "CTRLSRN=%s) — the template's TILT is left unchanged."
                % (m_no.group(1), sector_id, (pos + 1) if pos is not None else "?", srn)
            )
        deviceno_tilt[m_no.group(1)] = tilt
        report.append({
            "deviceno": m_no.group(1), "ctrlsrn": srn, "sector": sector_id,
            "device_pos": pos, "tilt": tilt, "matched_by": matched_by,
            "serial": None, "devicename": None,
        })
    report_by_no = {r["deviceno"]: r for r in report}

    # Pass 2: rewrite lines in place, preserving everything else verbatim.
    out_lines = []
    for ln in tmpl_lines:
        stripped = ln.lstrip()
        if stripped.startswith(add_prefix):
            m_no = _RE_DEVICENO.search(ln)
            rec = report_by_no.get(m_no.group(1)) if m_no else None

            def _sub_devicename(m, rec=rec):
                tokens = m.group(1).split("_")
                suffix = "_".join(tokens[prefix_tokens:])
                new = new_prefix + "_" + suffix if suffix else new_prefix
                if rec is not None:
                    rec["devicename"] = new
                return 'DEVICENAME="%s"' % new

            ln = _RE_DEVICENAME.sub(_sub_devicename, ln)

            m_srn = _RE_CTRLSRN.search(ln)
            if m_srn:
                srn = m_srn.group(1)

                def _sub_serial(m, srn=srn, rec=rec):
                    suffix = m.group(1)[-suf_len:]
                    new = serials.get((srn, suffix))
                    if new is None:
                        warnings.append(
                            "No RET-input serial for CTRLSRN=%s ending %r — the "
                            "template's SERIALNO %s is left in place."
                            % (srn, suffix, m.group(1))
                        )
                        return m.group(0)
                    used_serials.add((srn, suffix))
                    if rec is not None:
                        rec["serial"] = new
                    return 'SERIALNO="%s"' % new

                ln = _RE_SERIALNO.sub(_sub_serial, ln)
            out_lines.append(ln)
        elif stripped.startswith(tilt_prefixes):
            m_no = _RE_DEVICENO.search(ln)
            if m_no:
                tilt = deviceno_tilt.get(m_no.group(1))
                if tilt is not None:
                    ln = _RE_TILT.sub("TILT=%d" % int(round(float(tilt))), ln)
                elif m_no.group(1) not in deviceno_tilt:
                    warnings.append(
                        "%s DEVICENO=%s has no matching %s line in the template."
                        % (stripped.split(":", 1)[0], m_no.group(1), add_prefix)
                    )
            out_lines.append(ln)
        else:
            out_lines.append(ln)

    # Serials pasted in but never placed: usually the wrong site or a template
    # that covers fewer sectors than the site actually has.
    unused = sorted(set(serials) - used_serials)
    if unused:
        warnings.append(
            "%d RET-input serial(s) were not used: %s"
            % (len(unused), ", ".join("CTRLSRN=%s %s" % (s, serials[(s, x)])
                                      for s, x in unused[:12]))
            + (" …" if len(unused) > 12 else "")
        )

    return "".join(out_lines), warnings, report


def format_text_report(report, site=None):
    """Render the per-device mapping table shown next to the MML preview."""
    lines = []
    if site:
        lines.append("Site: %s" % site)
    lines.append("%-8s %-8s %-7s %-7s %-8s %s"
                 % ("DEVNO", "CTRLSRN", "SECTOR", "DEV#", "TILT", "DEVICENAME / SERIALNO"))
    lines.append("-" * 86)
    for r in report:
        lines.append("%-8s %-8s %-7s %-7s %-8s %s"
                     % (r["deviceno"], r["ctrlsrn"] or "-", r["sector"] or "?",
                        (r["device_pos"] + 1) if r["device_pos"] is not None else "?",
                        "-" if r["tilt"] is None else r["tilt"],
                        "%s  %s" % (r["devicename"] or "-", r["serial"] or "(unchanged)")))
    return "\n".join(lines)


def write_text_output(template_path, input_path, cdd_path, sheet, mapping, output_path,
                      input_text=None, clusters=None):
    """build_text_output + write to ``output_path``. Returns warnings."""
    text, warnings, _report = build_text_output(
        template_path, input_path, cdd_path, sheet, mapping,
        input_text=input_text, clusters=clusters,
    )
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(text)
    return warnings
