"""Fill the reconciliation template by editing its .xlsx package directly.

Only the data sheets' rows, their tables' ranges, the report header cells and the
workbook calculation flags change. Every other part is copied byte for byte, so
pivot tables, PIVOTBY reports, dynamic-array metadata and formatting survive
(openpyxl would drop several of these on save).

Row 2 of each data sheet is the *prototype row*: it supplies each column's style
and, for formula columns, the formula cell to repeat on every row.
"""
from __future__ import annotations

import io
import math
import posixpath
import re
import zipfile
from datetime import date, datetime
from html import unescape
from xml.sax.saxutils import escape

from .normalise import excel_serial

FORMULA_TABLES = {"reco_datatable", "BooksCurrent", "BooksPrevNext", "GST2bCurrent", "GST2bPrev"}
REPORT_SHEETS = (
    "Recon Summary", "Recon - To be Reclaimed", "Recon - ITC not to be availed",
    "Recon – Tax overview", "Recon - To Follow up",
)

_ILLEGAL_XML = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")
_CELL = re.compile(r"<c\b([^>]*?)(/>|>(.*?)</c>)", re.S)
_ROW = re.compile(r"<row\b([^>]*?)(/>|>(.*?)</row>)", re.S)
_ATTR = re.compile(r'([\w:]+)="([^"]*)"')


class TemplateError(Exception):
    """The template can't be used. The message is safe to print."""


# --------------------------------------------------------------------------- helpers

def col_index(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n


def col_letters(n: int) -> str:
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _attrs(text: str) -> dict:
    return dict(_ATTR.findall(text))


def _attr_text(attrs: dict) -> str:
    return "".join(f' {k}="{v}"' for k, v in attrs.items())


def _ref_col(ref: str) -> str:
    return re.match(r"[A-Z]+", ref).group(0)


def _resolve(base_part: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(base_part), target))


def _rels_path(part: str) -> str:
    return posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")


def _relationships(parts: dict, part: str) -> list:
    data = parts.get(_rels_path(part))
    if data is None:
        return []
    out = []
    for m in re.finditer(r"<Relationship\b([^>]*)/?>", data.decode("utf-8")):
        a = _attrs(m.group(1))
        out.append((a.get("Id"), a.get("Type", ""), _resolve(part, a.get("Target", ""))))
    return out


def _shared_strings(parts: dict) -> list:
    data = parts.get("xl/sharedStrings.xml")
    if data is None:
        return []
    out = []
    for si in re.finditer(r"<si>(.*?)</si>", data.decode("utf-8"), re.S):
        out.append(unescape("".join(re.findall(r"<t\b[^>]*>(.*?)</t>", si.group(1), re.S))))
    return out


def _cell_text(attrs: dict, inner: str, shared: list) -> str:
    if attrs.get("t") == "s":
        m = re.search(r"<v>(\d+)</v>", inner or "")
        return shared[int(m.group(1))] if m else ""
    if attrs.get("t") == "inlineStr":
        return unescape("".join(re.findall(r"<t\b[^>]*>(.*?)</t>", inner or "", re.S)))
    m = re.search(r"<v>(.*?)</v>", inner or "", re.S)
    return unescape(m.group(1)) if m else ""


def _value_cell(ref: str, style, value) -> str:
    s = f' s="{style}"' if style is not None else ""
    if value is None or (isinstance(value, float) and (math.isnan(value) or math.isinf(value))):
        return f'<c r="{ref}"{s}/>'
    if isinstance(value, (date, datetime)):
        return f'<c r="{ref}"{s}><v>{excel_serial(value)}</v></c>'
    if isinstance(value, bool):
        value = str(value)
    if isinstance(value, (int, float)):
        return f'<c r="{ref}"{s}><v>{value!r}</v></c>'
    text = escape(_ILLEGAL_XML.sub("", str(value)))
    return f'<c r="{ref}"{s} t="inlineStr"><is><t xml:space="preserve">{text}</t></is></c>'


class _Proto:
    """Prototype of one column, taken from template row 2."""

    def __init__(self, attrs: dict, inner: str):
        self.style = attrs.get("s")
        self.formula = None
        m = re.search(r"<f\b([^>]*?)(/>|>(.*?)</f>)", inner or "", re.S)
        if m:
            fattrs = _attrs(m.group(1))
            if fattrs.get("t") == "shared":  # keep the text, drop the sharing
                fattrs = {}
            self.formula = (fattrs, m.group(3) or "")
            self.cell_attrs = {k: v for k, v in attrs.items() if k not in ("r", "t")}

    def cell(self, ref: str, value) -> str:
        if self.formula is None:
            return _value_cell(ref, self.style, value)
        fattrs, ftext = self.formula
        fattrs = dict(fattrs)
        if "ref" in fattrs:
            fattrs["ref"] = ref
        return f'<c r="{ref}"{_attr_text(self.cell_attrs)}><f{_attr_text(fattrs)}>{ftext}</f></c>'


# --------------------------------------------------------------------------- sheet edit

def _fill_sheet(xml: str, rows: list, shared: list, table: str, sheet: str):
    m = re.search(r"<sheetData\s*/>|<sheetData>(.*?)</sheetData>", xml, re.S)
    if not m:
        raise TemplateError(f"sheet {sheet} has no data")
    body = m.group(1) or ""
    tpl_rows = {int(_attrs(r.group(1)).get("r", 0)): r for r in _ROW.finditer(body)}
    if 1 not in tpl_rows:
        raise TemplateError(f"sheet {sheet} has no header row")

    header_row = tpl_rows[1]
    columns = {}  # header text -> column index
    for c in _CELL.finditer(header_row.group(3) or ""):
        a = _attrs(c.group(1))
        text = _cell_text(a, c.group(3), shared).strip()
        if text:
            columns[text] = col_index(_ref_col(a["r"]))
    if not columns:
        raise TemplateError(f"sheet {sheet} has no column headers")

    protos, row_attrs = {}, {}
    if 2 in tpl_rows:
        proto_row = tpl_rows[2]
        row_attrs = {k: v for k, v in _attrs(proto_row.group(1)).items() if k != "r"}
        for c in _CELL.finditer(proto_row.group(3) or ""):
            a = _attrs(c.group(1))
            protos[col_index(_ref_col(a["r"]))] = _Proto(a, c.group(3))
    has_formula = any(p.formula for p in protos.values())
    if table in FORMULA_TABLES and not has_formula:
        raise TemplateError(f"template missing formulas: {table}")

    unknown = {k for r in rows for k in r} - set(columns)
    if unknown:
        raise TemplateError(f"sheet {sheet} has no column(s): {', '.join(sorted(unknown))[:200]}")

    last_col = max(list(columns.values()) + list(protos))
    order = sorted(set(columns.values()) | set(protos))
    by_index = {i: h for h, i in columns.items()}
    out_rows = [header_row.group(0)]
    data = rows if rows else [{}]
    for n, values in enumerate(data, 2):
        cells = []
        for ci in order:
            ref = f"{col_letters(ci)}{n}"
            proto = protos.get(ci)
            value = values.get(by_index.get(ci))
            if proto is not None:
                cells.append(proto.cell(ref, value))
            elif value is not None:
                cells.append(_value_cell(ref, None, value))
        attrs = dict(row_attrs, r=str(n))
        attrs.pop("spans", None)
        out_rows.append(f"<row{_attr_text(attrs)}>{''.join(cells)}</row>")

    last_row = len(data) + 1
    new_xml = xml[:m.start()] + "<sheetData>" + "".join(out_rows) + "</sheetData>" + xml[m.end():]
    new_xml = re.sub(r'(<dimension ref=")([A-Z]+\d+)(?::[A-Z]+\d+)?(")',
                     lambda d: f"{d.group(1)}{d.group(2)}:{col_letters(last_col)}{last_row}{d.group(3)}",
                     new_xml, count=1)
    return new_xml, last_row


def _resize_table(xml: str, last_row: int) -> str:
    def fix(m):
        start, end_col = m.group(2), m.group(3)
        return f'{m.group(1)}{start}:{end_col}{last_row}"'
    xml = re.sub(r'(<table\b[^>]*?\sref=")([A-Z]+\d+):([A-Z]+)\d+"', fix, xml, count=1)
    return re.sub(r'(<autoFilter\b[^>]*?\sref=")([A-Z]+\d+):([A-Z]+)\d+"', fix, xml, count=1)


def _set_cell(xml: str, ref: str, value) -> str:
    pattern = re.compile(r'<c\b(?=[^>]*\br="%s")([^>]*?)(/>|>(.*?)</c>)' % ref, re.S)
    m = pattern.search(xml)
    if not m:
        return xml
    style = _attrs(m.group(1)).get("s")
    return xml[:m.start()] + _value_cell(ref, style, value) + xml[m.end():]


# --------------------------------------------------------------------------- package

def fill_template(template: bytes, sheets: dict, header: dict) -> bytes:
    try:
        src = zipfile.ZipFile(io.BytesIO(template))
    except zipfile.BadZipFile:
        raise TemplateError("template is not an .xlsx file") from None
    infos = src.infolist()
    parts = {i.filename: src.read(i.filename) for i in infos}
    if "xl/workbook.xml" not in parts:
        raise TemplateError("template is not an .xlsx file")

    wb_xml = parts["xl/workbook.xml"].decode("utf-8")
    rels = {rid: target for rid, _, target in _relationships(parts, "xl/workbook.xml")}
    sheet_parts = {}
    for m in re.finditer(r"<sheet\b([^>]*)/>", wb_xml):
        a = _attrs(m.group(1))
        rid = a.get("r:id") or next((v for k, v in a.items() if k.endswith(":id")), None)
        sheet_parts[unescape(a.get("name", ""))] = rels.get(rid)

    shared = _shared_strings(parts)
    for name, rows in sheets.items():
        part = sheet_parts.get(name)
        if not part or part not in parts:
            raise TemplateError(f"template has no sheet {name}")
        tables = [t for _, typ, t in _relationships(parts, part) if typ.endswith("/table")]
        table_part = tables[0] if tables else None
        table_name = ""
        if table_part:
            tm = re.search(r'<table\b[^>]*\bname="([^"]+)"', parts[table_part].decode("utf-8"))
            table_name = tm.group(1) if tm else ""
        xml, last_row = _fill_sheet(parts[part].decode("utf-8"), rows, shared, table_name, name)
        parts[part] = xml.encode("utf-8")
        if table_part:
            parts[table_part] = _resize_table(parts[table_part].decode("utf-8"), last_row).encode("utf-8")

    if header:
        for name in REPORT_SHEETS:
            part = sheet_parts.get(name)
            if part in parts:
                xml = parts[part].decode("utf-8")
                for ref, value in header.items():
                    xml = _set_cell(xml, ref, value)
                parts[part] = xml.encode("utf-8")

    # recalculate everything when opened; the calc chain would be stale
    if re.search(r"<calcPr\b", wb_xml):
        if "fullCalcOnLoad" not in wb_xml:
            wb_xml = re.sub(r"<calcPr\b", '<calcPr fullCalcOnLoad="1"', wb_xml, count=1)
        else:
            wb_xml = re.sub(r'fullCalcOnLoad="[^"]*"', 'fullCalcOnLoad="1"', wb_xml)
    else:
        anchor = "</definedNames>" if "</definedNames>" in wb_xml else "</sheets>"
        wb_xml = wb_xml.replace(anchor, anchor + '<calcPr fullCalcOnLoad="1"/>', 1)
    parts["xl/workbook.xml"] = wb_xml.encode("utf-8")
    parts.pop("xl/calcChain.xml", None)
    wb_rels = "xl/_rels/workbook.xml.rels"
    parts[wb_rels] = re.sub(rb"<Relationship\b[^>]*calcChain[^>]*/>", b"", parts[wb_rels])
    parts["[Content_Types].xml"] = re.sub(
        rb'<Override\b[^>]*PartName="/xl/calcChain\.xml"[^>]*/>', b"", parts["[Content_Types].xml"])

    for name in list(parts):
        if re.match(r"xl/pivotCache/pivotCacheDefinition\d+\.xml$", name):
            x = parts[name].decode("utf-8")
            if not re.search(r'refreshOnLoad="(1|true)"', x):
                x = re.sub(r'\s*refreshOnLoad="[^"]*"', "", x)
                x = re.sub(r"<pivotCacheDefinition\b", '<pivotCacheDefinition refreshOnLoad="1"', x, count=1)
                parts[name] = x.encode("utf-8")

    out = io.BytesIO()
    with zipfile.ZipFile(out, "w") as z:
        for info in infos:
            if info.filename in parts:
                z.writestr(info, parts[info.filename])
    return out.getvalue()
