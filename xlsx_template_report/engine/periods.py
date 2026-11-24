# Part of Excel Template Reports. License LGPL-3, see LICENSE file. Copyright (c) 2026 Gray Bob.
"""Relative reporting periods ("last week", "month to date", ...).

All dates are local calendar dates in the report's timezone; ``date_to`` is
inclusive. The caller decides what "today" is (already converted to the
report timezone).
"""
from collections import namedtuple
from datetime import date, timedelta

Period = namedtuple("Period", ["date_from", "date_to", "label"])

PERIOD_TYPES = [
    ("today", "Today"),
    ("yesterday", "Yesterday"),
    ("last_n_days", "Last N days (excluding today)"),
    ("this_week", "This week to date"),
    ("last_week", "Last week"),
    ("this_month", "This month to date"),
    ("last_month", "Last month"),
    ("this_quarter", "This quarter to date"),
    ("last_quarter", "Last quarter"),
    ("this_year", "This year to date"),
    ("last_year", "Last year"),
]
_VALID = {key for key, _label in PERIOD_TYPES}


def _month_start(d):
    return d.replace(day=1)


def _add_months(d, months):
    """First day of the month ``months`` away from ``d``'s month."""
    index = d.year * 12 + (d.month - 1) + months
    return date(index // 12, index % 12 + 1, 1)


def _quarter_start(d):
    return date(d.year, 3 * ((d.month - 1) // 3) + 1, 1)


def compute_period(period_type, today, week_start=0, n_days=7):
    """Return a :class:`Period` for ``period_type`` relative to ``today``.

    :param week_start: first day of the week, Monday=0 ... Sunday=6
    :param n_days: only used by ``last_n_days``
    """
    if period_type not in _VALID:
        raise ValueError("Unknown period type: %r" % (period_type,))
    if not 0 <= int(week_start) <= 6:
        raise ValueError("week_start must be between 0 and 6")
    week_start = int(week_start)

    if period_type == "today":
        start = end = today
    elif period_type == "yesterday":
        start = end = today - timedelta(days=1)
    elif period_type == "last_n_days":
        if n_days < 1:
            raise ValueError("n_days must be >= 1")
        end = today - timedelta(days=1)
        start = today - timedelta(days=n_days)
    elif period_type in ("this_week", "last_week"):
        this_week_start = today - timedelta(days=(today.weekday() - week_start) % 7)
        if period_type == "this_week":
            start, end = this_week_start, today
        else:
            start = this_week_start - timedelta(days=7)
            end = this_week_start - timedelta(days=1)
    elif period_type == "this_month":
        start, end = _month_start(today), today
    elif period_type == "last_month":
        start = _add_months(today, -1)
        end = _month_start(today) - timedelta(days=1)
    elif period_type == "this_quarter":
        start, end = _quarter_start(today), today
    elif period_type == "last_quarter":
        current = _quarter_start(today)
        start = _add_months(current, -3)
        end = current - timedelta(days=1)
    elif period_type == "this_year":
        start, end = date(today.year, 1, 1), today
    else:  # last_year
        start, end = date(today.year - 1, 1, 1), date(today.year - 1, 12, 31)

    label = start.isoformat() if start == end else "%s → %s" % (start.isoformat(), end.isoformat())
    return Period(start, end, label)
