"""Local wall-clock times in the mailbox's time zone, the only form the model and the CEO use."""

from __future__ import annotations

import re
from datetime import UTC, date, datetime, timedelta
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

from .graph import GraphError

DAYS_IT = ("lun", "mar", "mer", "gio", "ven", "sab", "dom")


def zone(name: str) -> ZoneInfo:
    try:
        return ZoneInfo(name or "Europe/Rome")
    except ZoneInfoNotFoundError:
        return ZoneInfo("Europe/Rome")


def parse_local(value: str, tz: ZoneInfo, *, field: str, end_of_day: bool = False) -> datetime:
    """'2026-10-15T10:00', '2026-10-15 10:00', '2026-10-15' or 'today'/'oggi'/'tomorrow'/'domani' -> aware datetime."""
    text = str(value or "").strip().lower()
    today = datetime.now(tz).date()
    if text in ("today", "oggi"):
        text = today.isoformat()
    elif text in ("tomorrow", "domani"):
        text = (today + timedelta(days=1)).isoformat()
    m = re.fullmatch(r"(\d{4}-\d{2}-\d{2})(?:[t ](\d{1,2}):(\d{2})(?::\d{2}(?:\.\d+)?)?)?", text)
    if not m:
        raise GraphError("bad_time", f"{field} must look like 2026-10-15T10:00 (local time), got {value!r}")
    day = date.fromisoformat(m.group(1))
    if m.group(2) is None:
        return datetime(day.year, day.month, day.day, 23, 59, tzinfo=tz) if end_of_day else datetime(day.year, day.month, day.day, tzinfo=tz)
    return datetime(day.year, day.month, day.day, int(m.group(2)), int(m.group(3)), tzinfo=tz)


def graph_local(dt: datetime) -> str:
    """Graph dateTimeTimeZone.dateTime: wall clock without offset (the zone goes in timeZone)."""
    return dt.strftime("%Y-%m-%dT%H:%M:%S")


def utc_iso(dt: datetime) -> str:
    return dt.astimezone(UTC).strftime("%Y-%m-%dT%H:%M:%SZ")


def from_graph(value: dict | None, tz: ZoneInfo) -> datetime | None:
    """dateTimeTimeZone (already in the Prefer zone, or UTC) -> aware datetime in ``tz``."""
    if not isinstance(value, dict) or not value.get("dateTime"):
        return None
    raw = str(value["dateTime"])[:19]
    try:
        naive = datetime.fromisoformat(raw)
    except ValueError:
        return None
    # Every read sends Prefer: outlook.timezone=<mailbox zone>, so anything not UTC is already local.
    src = str(value.get("timeZone") or "UTC").upper()
    return naive.replace(tzinfo=UTC if src in ("UTC", "GMT", "COORDINATED UNIVERSAL TIME") else tz).astimezone(tz)


def human(dt: datetime | None, *, with_day: bool = True) -> str:
    if dt is None:
        return "?"
    return (f"{DAYS_IT[dt.weekday()]} {dt:%d/%m} " if with_day else "") + f"{dt:%H:%M}"


def human_range(start: datetime, end: datetime) -> str:
    same_day = start.date() == end.date()
    return f"{human(start)}–{human(end, with_day=not same_day)}"
