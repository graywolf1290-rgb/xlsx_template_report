# Part of Excel Template Reports. License LGPL-3, see LICENSE file. Copyright (c) 2026 Gray Bob.
"""Fill an .xlsx template with data.

Template model (see the product plan, section 3.4):

* **Data sheets** – one per data source, named exactly like the source. The
  engine clears the sheet, writes a header row + data rows and (re)creates an
  Excel Table with the same name, so formulas such as
  ``=SUM(sales[amount_total])`` keep working whatever the row count.
  The style of row 1 (header) and row 2 (first data row) of the template is
  reused for the generated rows.
* **Presentation sheets** – everything else. Only ``{{ placeholder }}`` cells
  are replaced; rows are never inserted or moved, so user formulas stay valid.

Known limitation: openpyxl drops charts and images when it re-saves a file.
:func:`inspect_template` reports this so the UI can warn at upload time.
"""
import io
import re
import zipfile
from collections import namedtuple
from copy import copy
from datetime import datetime

from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.utils import get_column_letter
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.worksheet.table import Table, TableStyleInfo

TableData = namedtuple("TableData", ["headers", "rows"])
RenderResult = namedtuple("RenderResult", ["content", "warnings"])
TemplateInfo = namedtuple("TemplateInfo", ["sheet_names", "table_names", "placeholders", "warnings"])

PLACEHOLDER_RE = re.compile(r"\{\{\s*([A-Za-z_][A-Za-z0-9_]*)\s*\}\}")
TABLE_NAME_RE = re.compile(r"^[A-Za-z_][A-Za-z0-9_]{0,30}$")
EXCEL_MAX_ROWS = 1048576
DEFAULT_TABLE_STYLE = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)


class TemplateError(ValueError):
    """The template file cannot be used."""


def valid_table_name(name):
    """Data source names double as sheet and Excel Table names."""
    return bool(name) and bool(TABLE_NAME_RE.match(name)) and not re.match(
        r"^[A-Za-z]{1,3}[0-9]+$", name)  # would look like a cell reference (A1, XFD10)


def _load(template):
    try:
        return load_workbook(io.BytesIO(template))
    except (InvalidFileException, zipfile.BadZipFile, KeyError, OSError) as exc:
        raise TemplateError("Not a valid .xlsx file: %s" % exc) from exc


def inspect_template(template):
    """Describe a template and list what will not survive rendering."""
    warnings = []
    try:
        with zipfile.ZipFile(io.BytesIO(template)) as archive:
            names = archive.namelist()
    except zipfile.BadZipFile as exc:
        raise TemplateError("Not a valid .xlsx file") from exc
    if any(n.startswith("xl/charts/") for n in names):
        warnings.append("The template contains charts. They will be removed from generated "
                        "files (use PivotTables or conditional formatting instead).")
    if any(n.startswith("xl/media/") for n in names):
        warnings.append("The template contains images. They will be removed from generated files.")
    if "xl/vbaProject.bin" in names:
        warnings.append("Macros are not supported and will be removed.")

    wb = _load(template)
    tables, placeholders = [], set()
    for ws in wb.worksheets:
        tables.extend(ws.tables.keys())
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and "{{" in cell.value:
                    placeholders.update(PLACEHOLDER_RE.findall(cell.value))
    return TemplateInfo(wb.sheetnames, sorted(tables), sorted(placeholders), warnings)


def _clean(value):
    if isinstance(value, str):
        return ILLEGAL_CHARACTERS_RE.sub("", value)
    if isinstance(value, datetime) and value.tzinfo is not None:
        # openpyxl refuses tz-aware datetimes; callers convert to local time first.
        return value.replace(tzinfo=None)
    if isinstance(value, (list, tuple, set, dict)):
        return str(value)
    return value


def _set_value(cell, value):
    """Write data as data: a text starting with '=' must never become a formula
    (formula injection, e.g. a partner named '=HYPERLINK(...)')."""
    value = _clean(value)
    cell.value = value
    if isinstance(value, str) and value.startswith("="):
        cell.data_type = "s"


def _row_styles(ws, row_idx):
    return {
        cell.column: copy(cell._style)
        for cell in ws[row_idx]
        if cell.has_style
    } if ws.max_row >= row_idx else {}


def _fill_data_sheet(wb, name, data):
    if not valid_table_name(name):
        raise TemplateError("Invalid data source name %r: use letters, digits and _ "
                            "(max 31 characters, not starting with a digit)." % name)
    if len(data.rows) + 1 > EXCEL_MAX_ROWS:
        raise TemplateError("Data source %r has too many rows for Excel." % name)
    headers = [str(h) for h in data.headers]
    if not headers:
        raise TemplateError("Data source %r has no columns." % name)
    if len(set(h.lower() for h in headers)) != len(headers):
        raise TemplateError("Data source %r has duplicate column names." % name)

    # Excel sheet and table names are case-insensitive and table names are
    # unique across the whole workbook.
    ws = next((s for s in wb.worksheets if s.title.lower() == name.lower()), None)
    for other in wb.worksheets:
        if other is not ws and any(t.lower() == name.lower() for t in other.tables):
            raise TemplateError("The table %r must be on a sheet named %r (found on sheet %r)."
                                % (name, name, other.title))
    if ws is None:
        ws = wb.create_sheet(name)
    header_styles = _row_styles(ws, 1)
    data_styles = _row_styles(ws, 2)
    style_info = DEFAULT_TABLE_STYLE
    for table_name in list(ws.tables.keys()):
        if table_name.lower() == name.lower():
            style_info = ws.tables[table_name].tableStyleInfo or DEFAULT_TABLE_STYLE
        del ws.tables[table_name]

    if ws.max_row:
        ws.delete_rows(1, ws.max_row)

    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col)
        _set_value(cell, header)
        if col in header_styles:
            cell._style = copy(header_styles[col])
    rows = data.rows or [[None] * len(headers)]  # an Excel Table needs one data row
    for r, row in enumerate(rows, start=2):
        for col, value in enumerate(row, start=1):
            cell = ws.cell(row=r, column=col)
            _set_value(cell, value)
            if col in data_styles:
                cell._style = copy(data_styles[col])

    ref = "A1:%s%d" % (get_column_letter(len(headers)), len(rows) + 1)
    ws.add_table(Table(displayName=name, ref=ref, tableStyleInfo=copy(style_info)))


def _replace_placeholders(ws, values, missing):
    for row in ws.iter_rows():
        for cell in row:
            text = cell.value
            if not isinstance(text, str) or "{{" not in text:
                continue
            full = PLACEHOLDER_RE.fullmatch(text.strip())
            if full:
                key = full.group(1)
                if key in values:
                    # Keep the native type (number, date) when the cell is only a placeholder.
                    _set_value(cell, values[key])
                else:
                    missing.add(key)
                continue

            def sub(match):
                key = match.group(1)
                if key not in values:
                    missing.add(key)
                    return match.group(0)
                value = values[key]
                return "" if value is None else str(value)

            _set_value(cell, PLACEHOLDER_RE.sub(sub, text))


def render(template, values, tables):
    """Render a report.

    :param template: bytes of an .xlsx file, or ``None`` to build a workbook
        that only contains the data sheets
    :param values: ``{placeholder: value}``
    :param tables: ``{source_name: TableData}``, written in the given order
    :return: :class:`RenderResult` (xlsx bytes, list of warning strings)
    """
    if template:
        wb = _load(template)
    else:
        wb = Workbook()
        wb.remove(wb.active)
        if not tables:
            raise TemplateError("Nothing to render: no template and no data source.")

    for name, data in tables.items():
        _fill_data_sheet(wb, name, data)

    missing = set()
    data_sheets = {name.lower() for name in tables}
    for ws in wb.worksheets:
        if ws.title.lower() not in data_sheets:
            _replace_placeholders(ws, values, missing)

    # openpyxl does not evaluate formulas: let Excel recompute everything on open.
    wb.calculation.fullCalcOnLoad = True
    for ws in wb.worksheets:
        for pivot in getattr(ws, "_pivots", []):
            pivot.cache.refreshOnLoad = True

    warnings = []
    if missing:
        warnings.append("Unknown placeholders left unchanged: %s" % ", ".join(sorted(missing)))

    out = io.BytesIO()
    wb.save(out)
    return RenderResult(fix_core_properties(out.getvalue()), warnings)


CORE_PROPERTIES = "docProps/core.xml"
_BAD_W3CDTF_RE = re.compile(rb"(\d{2}:\d{2}:\d{2}(?:\.\d+)?)\+00:00Z<")


def fix_core_properties(content):
    """Repair dates in docProps/core.xml written by openpyxl 3.1.2 on Python 3.12.

    That version (shipped with Odoo 18-20) stores timezone-aware dates and then
    appends "Z": ``2026-11-23T15:50:56+00:00Z``. Excel rejects it and opens the
    file in repair mode ("[Repaired]"). Fixed upstream in openpyxl 3.1.3; we
    patch the bytes so every openpyxl version produces a clean file.
    """
    with zipfile.ZipFile(io.BytesIO(content)) as archive:
        if CORE_PROPERTIES not in archive.namelist():
            return content
        core = archive.read(CORE_PROPERTIES)
        fixed = _BAD_W3CDTF_RE.sub(rb"\1Z<", core)
        if fixed == core:
            return content
        out = io.BytesIO()
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as new:
            for info in archive.infolist():
                data = fixed if info.filename == CORE_PROPERTIES else archive.read(info.filename)
                new.writestr(info, data, compress_type=zipfile.ZIP_DEFLATED)
    return out.getvalue()
