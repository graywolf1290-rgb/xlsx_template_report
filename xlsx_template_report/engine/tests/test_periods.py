# Part of Excel Template Reports. License LGPL-3, see LICENSE file. Copyright (c) 2026 Gray Bob.
from datetime import date

import pytest

from engine.periods import compute_period

# 2026-11-23 is a Monday.
MON = date(2026, 11, 23)
WED = date(2026, 11, 25)


@pytest.mark.parametrize("ptype,today,expected", [
    ("today", WED, (WED, WED)),
    ("yesterday", WED, (date(2026, 11, 24), date(2026, 11, 24))),
    ("yesterday", date(2026, 3, 1), (date(2026, 2, 28), date(2026, 2, 28))),
    ("this_week", WED, (MON, WED)),
    ("this_week", MON, (MON, MON)),
    ("last_week", WED, (date(2026, 11, 16), date(2026, 11, 22))),
    ("last_week", MON, (date(2026, 11, 16), date(2026, 11, 22))),
    ("this_month", WED, (date(2026, 11, 1), WED)),
    ("last_month", WED, (date(2026, 10, 1), date(2026, 10, 31))),
    ("last_month", date(2026, 1, 15), (date(2025, 12, 1), date(2025, 12, 31))),
    ("last_month", date(2024, 3, 31), (date(2024, 2, 1), date(2024, 2, 29))),
    ("this_quarter", WED, (date(2026, 10, 1), WED)),
    ("last_quarter", WED, (date(2026, 7, 1), date(2026, 9, 30))),
    ("last_quarter", date(2026, 2, 10), (date(2025, 10, 1), date(2025, 12, 31))),
    ("this_year", WED, (date(2026, 1, 1), WED)),
    ("last_year", WED, (date(2025, 1, 1), date(2025, 12, 31))),
])
def test_periods(ptype, today, expected):
    period = compute_period(ptype, today)
    assert (period.date_from, period.date_to) == expected


def test_week_start_sunday():
    # Wednesday 25th, weeks start Sunday -> this week started Sunday 22nd.
    period = compute_period("last_week", WED, week_start=6)
    assert (period.date_from, period.date_to) == (date(2026, 11, 15), date(2026, 11, 21))


def test_last_n_days_excludes_today():
    period = compute_period("last_n_days", WED, n_days=7)
    assert (period.date_from, period.date_to) == (date(2026, 11, 18), date(2026, 11, 24))


def test_label():
    assert compute_period("today", WED).label == "2026-11-25"
    assert compute_period("last_month", WED).label == "2026-10-01 → 2026-10-31"


@pytest.mark.parametrize("kwargs", [
    {"period_type": "nope"},
    {"period_type": "last_week", "week_start": 7},
    {"period_type": "last_n_days", "n_days": 0},
])
def test_invalid(kwargs):
    with pytest.raises(ValueError):
        compute_period(today=WED, **kwargs)
