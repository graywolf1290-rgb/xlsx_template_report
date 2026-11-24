# Part of Excel Template Reports. License LGPL-3, see LICENSE file. Copyright (c) 2026 Gray Bob.
"""Pure-Python core of the module.

Nothing in this package may import ``odoo``: it must stay testable with plain
pytest and reusable unchanged across Odoo versions. Only ``openpyxl`` and
``pytz`` (both shipped with Odoo) are allowed as third-party imports.
"""
from .periods import PERIOD_TYPES, Period, compute_period
from .render import (
    RenderResult,
    TableData,
    TemplateError,
    TemplateInfo,
    inspect_template,
    render,
    valid_table_name,
)
from .schedule import INTERVALS, next_occurrence

__all__ = [
    "INTERVALS",
    "TemplateError",
    "valid_table_name",
    "PERIOD_TYPES",
    "Period",
    "compute_period",
    "RenderResult",
    "TableData",
    "TemplateInfo",
    "inspect_template",
    "render",
    "next_occurrence",
]
