# Part of Excel Template Reports. License LGPL-3, see LICENSE file. Copyright (c) 2026 Gray Bob.
"""Next-run computation for report schedules.

Schedules are expressed in local wall-clock time of a timezone; the result is
a naive UTC datetime (the way Odoo stores Datetime fields).
"""
import calendar
from datetime import datetime, time, timedelta

import pytz

INTERVALS = ("daily", "weekly", "monthly")


def _matches(day, interval, weekday, monthday, last_day):
    if interval == "daily":
        return True
    if interval == "weekly":
        return day.weekday() == weekday
    days_in_month = calendar.monthrange(day.year, day.month)[1]
    if last_day:
        return day.day == days_in_month
    # Day 31 in a 30-day month (or 29-31 in February) runs on the last day.
    return day.day == min(monthday, days_in_month)


def next_occurrence(after_utc, interval, hour, minute, tz_name="UTC",
                    weekday=0, monthday=1, last_day=False):
    """Return the first scheduled moment strictly after ``after_utc``.

    :param after_utc: naive datetime in UTC
    :param interval: ``daily`` / ``weekly`` / ``monthly``
    :param weekday: Monday=0 ... Sunday=6 (weekly only)
    :param monthday: 1..31 (monthly only, clamped to the month length)
    :param last_day: monthly only, run on the last day of every month
    :return: naive datetime in UTC
    """
    if interval not in INTERVALS:
        raise ValueError("Unknown interval: %r" % (interval,))
    if not (0 <= hour <= 23 and 0 <= minute <= 59):
        raise ValueError("Invalid time %s:%s" % (hour, minute))
    if interval == "weekly" and not 0 <= weekday <= 6:
        raise ValueError("weekday must be between 0 and 6")
    if interval == "monthly" and not last_day and not 1 <= monthday <= 31:
        raise ValueError("monthday must be between 1 and 31")

    tz = pytz.timezone(tz_name or "UTC")
    after_local = pytz.utc.localize(after_utc).astimezone(tz)
    day = after_local.date()
    # 62 days covers any monthly schedule; leave margin for safety.
    for _i in range(400):
        if _matches(day, interval, weekday, monthday, last_day):
            # is_dst=None would raise on DST gaps/overlaps; False picks the
            # standard-time interpretation, which is good enough for reports.
            local_dt = tz.localize(datetime.combine(day, time(hour, minute)), is_dst=False)
            candidate = local_dt.astimezone(pytz.utc).replace(tzinfo=None)
            if candidate > after_utc:
                return candidate
        day += timedelta(days=1)
    raise RuntimeError("No occurrence found within 400 days")  # pragma: no cover
