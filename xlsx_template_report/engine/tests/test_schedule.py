# Part of Excel Template Reports. License LGPL-3, see LICENSE file. Copyright (c) 2026 Gray Bob.
from datetime import datetime

import pytest

from engine.schedule import next_occurrence

HCM = "Asia/Ho_Chi_Minh"  # UTC+7, no DST


def test_daily_later_today():
    # 2026-11-23 00:00 UTC = 07:00 local; 08:00 local today = 01:00 UTC
    assert next_occurrence(datetime(2026, 11, 23, 0, 0), "daily", 8, 0, HCM) == datetime(2026, 11, 23, 1, 0)


def test_daily_already_passed_goes_to_tomorrow():
    assert next_occurrence(datetime(2026, 11, 23, 1, 0), "daily", 8, 0, HCM) == datetime(2026, 11, 24, 1, 0)


def test_local_date_differs_from_utc_date():
    # 2026-11-23 20:00 UTC is already the 24th 03:00 in Ho Chi Minh City.
    assert next_occurrence(datetime(2026, 11, 23, 20, 0), "daily", 8, 0, HCM) == datetime(2026, 11, 24, 1, 0)


def test_weekly_monday_8am():
    # From Wednesday 25th -> next Monday 30th 08:00 local.
    result = next_occurrence(datetime(2026, 11, 25, 5, 0), "weekly", 8, 0, HCM, weekday=0)
    assert result == datetime(2026, 11, 30, 1, 0)


def test_monthly_31_clamps_to_month_end():
    result = next_occurrence(datetime(2026, 2, 1), "monthly", 8, 0, "UTC", monthday=31)
    assert result == datetime(2026, 2, 28, 8, 0)


def test_monthly_last_day():
    result = next_occurrence(datetime(2024, 2, 10), "monthly", 23, 30, "UTC", last_day=True)
    assert result == datetime(2024, 2, 29, 23, 30)


def test_dst_gap_does_not_crash():
    # 2026-03-08 02:30 does not exist in New York (clocks jump 02:00 -> 03:00).
    result = next_occurrence(datetime(2026, 3, 8, 0, 0), "daily", 2, 30, "America/New_York")
    assert result.date() == datetime(2026, 3, 8).date()


def test_missed_runs_collapse_to_one():
    # Server was down for a week: the next run is simply the next slot after "now".
    now = datetime(2026, 11, 30, 3, 0)
    assert next_occurrence(now, "daily", 8, 0, HCM) == datetime(2026, 12, 1, 1, 0)


@pytest.mark.parametrize("kwargs", [
    {"interval": "hourly"},
    {"interval": "daily", "hour": 24},
    {"interval": "weekly", "weekday": 7},
    {"interval": "monthly", "monthday": 0},
])
def test_invalid(kwargs):
    params = {"hour": 8, "minute": 0}
    params.update(kwargs)
    with pytest.raises(ValueError):
        next_occurrence(datetime(2026, 1, 1), **params)
