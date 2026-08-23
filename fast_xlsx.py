"""Fast, read-only .xlsx column reader used by ret_core to parse the CDD.

Why: the production CDD workbook is ~22 MB / 15 700 rows x 106 columns, and
``openpyxl``'s read-only ``iter_rows`` materialises every one of those ~1.6 M
cells — about 9 s per parse, paid again on every Preview/Generate click.

NewRET only ever needs ~8 columns of one sheet. This module therefore:

  1. resolves the header row generically (first ``max_scan`` rows only), then
  2. extracts *only the wanted columns* for the remaining rows with a single
     regex pass over the raw sheet XML.

That is ~0.5 s instead of ~9 s. ``read_columns`` returns exactly what a
generic reader would have returned for those columns, so ret_core can fall
back to openpyxl (``ret_core._read_sheet_openpyxl``) whenever anything about a
workbook is unusual.

Only the parts of SpreadsheetML NewRET actually meets are handled (shared
strings, inline strings, cached formula values, numbers, booleans). Anything
unexpected raises, which the caller turns into an openpyxl fallback.
"""
import re
import zipfile
from xml.etree.ElementTree import iterparse

_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_NS_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

_RE_ROW_SPLIT = re.compile(rb"<row[ >]")
_RE_ROW_NUM = re.compile(rb'<row[^>]*\br="(\d+)"')
_RE_CELL = re.compile(
    rb'<c r="([A-Z]+)\d+"([^>]*?)(?:/>|>(.*?)</c>)', re.DOTALL
)
_RE_V = re.compile(rb"<v[^>]*>(.*?)</v>", re.DOTALL)
_RE_T = re.compile(rb"<t[^>]*>(.*?)</t>", re.DOTALL)
_RE_TYPE = re.compile(rb't="([^"]*)"')


def col_letter(index0):
    """0-based column index -> spreadsheet column letters (0 -> 'A')."""
    n, out = index0 + 1, ""
    while n:
        n, rem = divmod(n - 1, 26)
        out = chr(65 + rem) + out
    return out


def col_index(letters):
    """Column letters -> 0-based index ('A' -> 0)."""
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def _unescape(b):
    s = b.decode("utf-8")
    if "&" not in s:
        return s
    return (s.replace("&lt;", "<").replace("&gt;", ">")
             .replace("&quot;", '"').replace("&apos;", "'")
             .replace("&amp;", "&"))


def _number(text):
    """XML numeric text -> int when integral, else float."""
    try:
        f = float(text)
    except ValueError:
        return text
    return int(f) if f.is_integer() else f


class Workbook:
    """A lazily-parsed .xlsx opened for column extraction."""

    def __init__(self, path):
        self.zf = zipfile.ZipFile(path)
        self._sst = None
        self._sheets = None

    def close(self):
        self.zf.close()

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        self.close()

    # ---------- workbook structure ----------
    @property
    def sheets(self):
        """Ordered {sheet name: zip part name}."""
        if self._sheets is None:
            rels = {}
            with self.zf.open("xl/_rels/workbook.xml.rels") as f:
                for _, el in iterparse(f, ("end",)):
                    if el.tag.endswith("}Relationship"):
                        target = el.get("Target", "")
                        target = target[3:] if target.startswith("../") else "xl/" + target.lstrip("/")
                        rels[el.get("Id")] = target
            sheets = {}
            with self.zf.open("xl/workbook.xml") as f:
                for _, el in iterparse(f, ("end",)):
                    if el.tag == _NS + "sheet":
                        part = rels.get(el.get(_NS_REL + "id"))
                        if part:
                            sheets[el.get("name")] = part
            self._sheets = sheets
        return self._sheets

    @property
    def sst(self):
        """Shared-string table (list of str)."""
        if self._sst is None:
            out = []
            try:
                f = self.zf.open("xl/sharedStrings.xml")
            except KeyError:
                self._sst = out
                return out
            with f:
                for _, el in iterparse(f, ("end",)):
                    if el.tag == _NS + "si":
                        out.append("".join(n.text or "" for n in el.iter(_NS + "t")))
                        el.clear()
            self._sst = out
        return self._sst

    # ---------- cell decoding ----------
    def _value(self, attrs, body):
        if body is None:
            return None
        m_t = _RE_TYPE.search(attrs)
        ctype = m_t.group(1) if m_t else b""
        if ctype == b"s":
            m = _RE_V.search(body)
            if m is None:
                return None
            try:
                return self.sst[int(m.group(1))]
            except (ValueError, IndexError):
                return None
        if ctype == b"inlineStr":
            parts = _RE_T.findall(body)
            return "".join(_unescape(p) for p in parts) if parts else None
        m = _RE_V.search(body)
        if m is None:
            return None
        raw = m.group(1)
        if not raw:
            return None
        if ctype in (b"str", b"e"):
            return _unescape(raw)
        if ctype == b"b":
            return raw == b"1"
        return _number(raw.decode("ascii", "replace"))

    # ---------- reading ----------
    def head_rows(self, sheet, max_rows):
        """First ``max_rows`` sheet rows as full tuples, like openpyxl would.

        Rows the file omits entirely (all-blank) are yielded as empty tuples so
        that a row's position here is its position in ``openpyxl.iter_rows``.
        """
        data = self.zf.read(self.sheets[sheet])
        rows = []
        for rnum, chunk in self._row_chunks(data):
            while len(rows) < rnum - 1:
                rows.append(())
            if len(rows) >= max_rows:
                break
            cells = {}
            for letters, attrs, body in _RE_CELL.findall(chunk):
                cells[col_index(letters.decode())] = self._value(attrs, body)
            width = (max(cells) + 1) if cells else 0
            rows.append(tuple(cells.get(i) for i in range(width)))
        return rows[:max_rows]

    def read_columns(self, sheet, indices, skip_rows=0):
        """Rows of ``sheet`` after ``skip_rows``, as {col index: value} dicts.

        Only the columns in ``indices`` (0-based) are decoded; every other cell
        in the sheet is skipped by the regex, which is where the speed-up comes
        from. Blank cells are simply absent from a row's dict; rows the file
        omits are not emitted at all (they carry no data by definition).
        """
        indices = sorted(set(indices))
        if not indices:
            return []
        letters = "|".join(col_letter(i) for i in indices)
        pat = re.compile(
            (r'<c r="(%s)\d+"([^>]*?)(?:/>|>(.*?)</c>)' % letters).encode(),
            re.DOTALL,
        )
        data = self.zf.read(self.sheets[sheet])
        out = []
        for rnum, chunk in self._row_chunks(data):
            if rnum <= skip_rows:
                continue
            row = {}
            for lt, attrs, body in pat.findall(chunk):
                v = self._value(attrs, body)
                if v is not None and v != "":
                    row[col_index(lt.decode())] = v
            if row:
                out.append(row)
        return out

    @staticmethod
    def _row_chunks(data):
        """Yield (1-based row number, raw XML of that <row>) in sheet order."""
        bounds = [(m.start(), m.end()) for m in _RE_ROW_SPLIT.finditer(data)]
        for i, (s, _e) in enumerate(bounds):
            end = bounds[i + 1][0] if i + 1 < len(bounds) else len(data)
            chunk = data[s:end]
            m = _RE_ROW_NUM.match(chunk)
            yield (int(m.group(1)) if m else i + 1), chunk
