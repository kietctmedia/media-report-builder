# -*- coding: utf-8 -*-
"""Lịch làm việc Việt Nam: đếm ngày công chuẩn (T2-T6) và trừ ngày lễ.

Bảng ngày lễ nằm ở config/holidays.json để user tự sửa — KHÔNG hardcode trong code,
vì lịch nghỉ bù / lịch âm mỗi năm mỗi khác và công ty có thể nghỉ khác luật.
"""
from __future__ import annotations

import datetime as _dt
import json
import os
from typing import Iterable

_CONFIG = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(
    os.path.abspath(__file__)))), 'config', 'holidays.json')


def _load_holidays() -> dict:
    try:
        with open(_CONFIG, encoding='utf-8') as fh:
            return json.load(fh)
    except Exception:
        return {}


def holidays_in_month(year: int, month: int) -> list[dict]:
    """Trả về [{date: 'YYYY-MM-DD', name: str}] các ngày lễ rơi vào T2-T6 của tháng."""
    data = _load_holidays()
    out = []
    for item in data.get(str(year), []):
        try:
            d = _dt.date.fromisoformat(item['date'])
        except Exception:
            continue
        if d.month == month and d.weekday() < 5:
            out.append({'date': item['date'], 'name': item.get('name', '')})
    return sorted(out, key=lambda x: x['date'])


def weekdays_in_month(year: int, month: int) -> int:
    """Số ngày T2-T6 trong tháng (chưa trừ lễ)."""
    if month == 12:
        nxt = _dt.date(year + 1, 1, 1)
    else:
        nxt = _dt.date(year, month + 1, 1)
    days = (nxt - _dt.date(year, month, 1)).days
    return sum(1 for d in range(1, days + 1)
               if _dt.date(year, month, d).weekday() < 5)


def suggest_cong_chuan(year: int, month: int) -> dict:
    """Gợi ý công chuẩn = số ngày T2-T6 trừ ngày lễ đã khai báo."""
    wd = weekdays_in_month(year, month)
    hol = holidays_in_month(year, month)
    return {
        'weekdays': wd,
        'holidays': hol,
        'suggested': wd - len(hol),
        'note': f'{wd} ngày T2–T6' + (f' − {len(hol)} ngày lễ' if hol else ''),
    }


def month_range(year: int, month: int, lookback: int = 1) -> list[tuple[int, int]]:
    """Tháng báo cáo + `lookback` tháng trước đó (để bắt task carryover)."""
    out = [(year, month)]
    y, m = year, month
    for _ in range(lookback):
        m -= 1
        if m == 0:
            m, y = 12, y - 1
        out.append((y, m))
    return out


def parse_dmy(value) -> _dt.datetime | None:
    """Base.vn xuất ngày dạng chuỗi 'DD/MM/YYYY HH:MM' (đôi khi thiếu giờ)."""
    s = str(value or '').strip()
    if not s:
        return None
    for fmt in ('%d/%m/%Y %H:%M:%S', '%d/%m/%Y %H:%M', '%d/%m/%Y'):
        try:
            return _dt.datetime.strptime(s, fmt)
        except ValueError:
            continue
    return None


def month_of(value) -> tuple[int, int] | None:
    d = parse_dmy(value)
    return (d.year, d.month) if d else None


def label(months: Iterable[tuple[int, int]]) -> str:
    return ', '.join(f'{m:02d}/{y}' for y, m in months)
