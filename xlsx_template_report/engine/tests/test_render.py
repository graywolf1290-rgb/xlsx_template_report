# Part of Excel Template Reports. License LGPL-3, see LICENSE file. Copyright (c) 2026 Gray Bob.
import io
from datetime import date, datetime, timedelta, timezone

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Font
from openpyxl.worksheet.table import Table, TableStyleInfo

import re
import zipfile

from engine.render import (
    TableData,
    TemplateError,
    fix_core_properties,
    inspect_template,
    render,
    valid_table_name,
)

W3CDTF_RE = re.compile(r"^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(\.\d+)?(Z|[+-]\d{2}:\d{2})$")


def _core_dates(content):
    core = zipfile.ZipFile(io.BytesIO(content)).read("docProps/core.xml").decode()
    return re.findall(r"<dcterms:(?:created|modified)[^>]*>([^<]*)<", core)


def test_core_properties_dates_are_valid_for_excel():
    content = render(None, {}, {"t": TableData(["a"], [[1]])}).content
    dates = _core_dates(content)
    assert dates and all(W3CDTF_RE.match(d) for d in dates), dates


def test_fix_core_properties_repairs_openpyxl_312_output():
    good = render(None, {}, {"t": TableData(["a"], [[1]])}).content
    src = zipfile.ZipFile(io.BytesIO(good))
    bad = io.BytesIO()
    with zipfile.ZipFile(bad, "w", zipfile.ZIP_DEFLATED) as out:
        for info in src.infolist():
            data = src.read(info.filename)
            if info.filename == "docProps/core.xml":
                # what openpyxl 3.1.2 on Python 3.12 writes
                data = re.sub(rb"(<dcterms:modified[^>]*>)[^<]*<", rb"\g<1>2026-11-23T15:50:56+00:00Z<", data)
            out.writestr(info, data)
    assert "2026-11-23T15:50:56+00:00Z" in _core_dates(bad.getvalue())
    fixed = fix_core_properties(bad.getvalue())
    assert "2026-11-23T15:50:56Z" in _core_dates(fixed)
    assert zipfile.ZipFile(io.BytesIO(fixed)).namelist() == src.namelist()
    load_workbook(io.BytesIO(fixed))  # still a valid workbook


def _to_bytes(wb):
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def _load(content):
    return load_workbook(io.BytesIO(content))


@pytest.fixture
def template():
    """Presentation sheet + a styled data sheet with an existing table."""
    wb = Workbook()
    summary = wb.active
    summary.title = "Summary"
    summary["A1"] = "Sales report {{ period_label }}"
    summary["A2"] = "{{date_from}}"
    summary["A3"] = "=SUM(sales[amount])"
    summary["A4"] = "{{ unknown_key }}"

    data = wb.create_sheet("sales")
    data["A1"], data["B1"] = "old_header", "amount"
    data["A1"].font = Font(bold=True)
    data["A2"], data["B2"] = "x", 1
    data["B2"].number_format = "#,##0.00"
    data.add_table(Table(displayName="sales", ref="A1:B2",
                         tableStyleInfo=TableStyleInfo(name="TableStyleLight9")))
    return _to_bytes(wb)


def test_render_fills_data_and_placeholders(template):
    rows = [["Alice", 10.5], ["Bob", 20], ["Carol", None]]
    result = render(template, {"period_label": "2026-11", "date_from": date(2026, 11, 1)},
                    {"sales": TableData(["salesperson", "amount"], rows)})
    wb = _load(result.content)

    summary = wb["Summary"]
    assert summary["A1"].value == "Sales report 2026-11"
    # A cell that is only a placeholder keeps the native type.
    assert summary["A2"].value.date() == date(2026, 11, 1)
    # Formulas referencing the table are untouched.
    assert summary["A3"].value == "=SUM(sales[amount])"
    assert summary["A4"].value == "{{ unknown_key }}"
    assert any("unknown_key" in w for w in result.warnings)

    data = wb["sales"]
    assert [c.value for c in data[1]] == ["salesperson", "amount"]
    assert data["A4"].value == "Carol"
    table = data.tables["sales"]
    assert table.ref == "A1:B4"
    assert table.tableStyleInfo.name == "TableStyleLight9"
    # Styles of the template rows are reused.
    assert data["A1"].font.bold
    assert data["B3"].number_format == "#,##0.00"
    assert wb.calculation.fullCalcOnLoad


def test_render_is_repeatable(template):
    first = render(template, {}, {"sales": TableData(["a", "b"], [[1, 2]] * 5)}).content
    second = render(first, {}, {"sales": TableData(["a", "b"], [[1, 2]])}).content
    data = _load(second)["sales"]
    assert data.max_row == 2
    assert data.tables["sales"].ref == "A1:B2"


def test_empty_source_still_has_valid_table(template):
    result = render(template, {}, {"sales": TableData(["a", "b"], [])})
    assert _load(result.content)["sales"].tables["sales"].ref == "A1:B2"


def test_no_template_builds_data_only_workbook():
    result = render(None, {}, {"partners": TableData(["name"], [["A"], ["B"]])})
    wb = _load(result.content)
    assert wb.sheetnames == ["partners"]
    assert wb["partners"]["A3"].value == "B"


def test_case_insensitive_sheet_match(template):
    result = render(template, {}, {"SALES": TableData(["a"], [[1]])})
    wb = _load(result.content)
    assert "SALES" not in wb.sheetnames  # reused the existing "sales" sheet
    assert wb["sales"]["A2"].value == 1


def test_value_cleaning():
    aware = datetime(2026, 1, 1, 8, 0, tzinfo=timezone(timedelta(hours=7)))
    result = render(None, {}, {"t": TableData(["a", "b", "c"], [["bad\x01char", aware, [1, 2]]])})
    ws = _load(result.content)["t"]
    assert ws["A2"].value == "badchar"
    assert ws["B2"].value == datetime(2026, 1, 1, 8, 0)
    assert ws["C2"].value == "[1, 2]"


def test_formula_injection_is_neutralised():
    evil = '=HYPERLINK("http://evil.example","click")'
    wb = Workbook()
    wb.active["A1"] = "{{ name }}"
    wb.active["B1"] = "=1+1"
    result = render(_to_bytes(wb), {"name": evil}, {"t": TableData(["name"], [[evil]])})
    out = _load(result.content)
    for cell in (out["t"]["A2"], out.worksheets[0]["A1"]):
        assert cell.data_type == "s"
        assert cell.value == evil
    # Formulas written by the template author stay formulas.
    assert out.worksheets[0]["B1"].data_type == "f"


@pytest.mark.parametrize("name,ok", [
    ("sales", True), ("_x", True), ("sales_2026", True),
    ("1sales", False), ("my sales", False), ("A1", False), ("XFD10", False),
    ("x" * 31, True), ("x" * 32, False), ("", False),
])
def test_valid_table_name(name, ok):
    assert valid_table_name(name) is ok


def test_errors():
    with pytest.raises(TemplateError):
        render(b"not a zip", {}, {})
    with pytest.raises(TemplateError):
        render(None, {}, {"bad name": TableData(["a"], [])})
    with pytest.raises(TemplateError):
        render(None, {}, {"t": TableData(["a", "A"], [])})
    with pytest.raises(TemplateError):
        render(None, {}, {})


def test_table_on_wrong_sheet_is_rejected():
    wb = Workbook()
    ws = wb.active
    ws.title = "Other"
    ws["A1"], ws["A2"] = "h", 1
    ws.add_table(Table(displayName="sales", ref="A1:A2"))
    with pytest.raises(TemplateError):
        render(_to_bytes(wb), {}, {"sales": TableData(["h"], [[1]])})


def test_inspect_template_warns_about_charts(template):
    wb = _load(template)
    chart = BarChart()
    chart.add_data(Reference(wb["sales"], min_col=2, min_row=1, max_row=2))
    wb["Summary"].add_chart(chart, "D2")
    info = inspect_template(_to_bytes(wb))
    assert info.sheet_names == ["Summary", "sales"]
    assert info.table_names == ["sales"]
    assert info.placeholders == ["date_from", "period_label", "unknown_key"]
    assert any("charts" in w for w in info.warnings)
