# -*- coding: utf-8 -*-
"""Lọc - dedup - phân loại - chấm điểm.

Luồng lọc (đúng như quy trình đã chạy tay từ T5 đến T9):
  1. Task đã được RATE (5 tiêu chí đều có điểm)
  2. Người làm task thuộc danh sách đánh giá
  3. Ngày tạo thuộc {tháng báo cáo} ∪ {N tháng trước} — lookback để bắt carryover
  4. Trừ toàn bộ ID đã nằm trong báo cáo tháng trước (dedup)
  5. Task tháng lookback thì chia 2 loại:
       - hoàn thành SAU khi báo cáo tháng trước chốt  -> carryover thật, mặc định LẤY
       - hoàn thành TRƯỚC đó                          -> đã từng bị loại tay, mặc định BỎ
"""
from __future__ import annotations

import datetime as _dt
import json
import os

from . import calendar_vn as cal
from . import rules as R

_TIERS_PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(
    os.path.abspath(__file__)))), 'config', 'kpi_tiers.json')


def load_tiers() -> dict:
    try:
        with open(_TIERS_PATH, encoding='utf-8') as fh:
            cfg = json.load(fh)
        return {'hours_per_day': float(cfg.get('gio_mot_ngay_cong', 8)),
                'tiers': sorted(cfg['bac'], key=lambda t: -float(t['min']))}
    except Exception:
        return {'hours_per_day': 8.0, 'tiers': [
            {'min': 8.9, 'ten': 'MỨC 3', 'tien': 3000000},
            {'min': 8.4, 'ten': 'MỨC 2', 'tien': 2500000},
            {'min': 8.0, 'ten': 'MỨC 1', 'tien': 2000000},
            {'min': 0, 'ten': 'Không đạt KPI', 'tien': 0}]}


def _num(v):
    if isinstance(v, (int, float)):
        return float(v) if v != '' else None
    s = str(v or '').strip()
    if not s:
        return None
    try:
        return float(s.replace(',', '.'))
    except ValueError:
        return None


def analyse(rows, year, month, persons, prev_ids, cutoff_iso=None, lookback=1):
    """Trả về (tasks, stats). tasks đã kèm đề xuất L/M/N và cờ include."""
    cfg = R.load_rate_card()
    months = cal.month_range(year, month, lookback)
    main_month, prev_months = months[0], set(months[1:])
    prev = set(prev_ids or [])
    cutoff = _dt.datetime.fromisoformat(cutoff_iso) if cutoff_iso else None
    pset = set(persons)

    stats = {'total_rows': len(rows), 'wrong_person': 0, 'out_of_month': 0,
             'not_rated': 0, 'text_rating': 0, 'duplicated': 0,
             'carryover_kept': 0, 'carryover_dropped': 0, 'no_time': 0}
    picked = []
    for r in rows:
        person = str(r.get('person') or '').strip()
        if person not in pset:
            stats['wrong_person'] += 1
            continue
        m = cal.month_of(r.get('ngaytao'))
        if m not in months:
            stats['out_of_month'] += 1
            continue
        raw = r.get('ratings') or []
        nums = [_num(x) for x in raw]
        if any(x is None for x in nums):
            if any(str(x).strip() for x in raw):
                stats['text_rating'] += 1          # widget đổi sang dạng chữ
            stats['not_rated'] += 1
            continue
        _id = str(r.get('id') or '').strip()
        if _id in prev:
            stats['duplicated'] += 1
            continue

        T = _num(r.get('tleader')) or 0.0
        if T <= 0:
            stats['no_time'] += 1

        is_carry = m in prev_months
        include, reason = True, ''
        if is_carry:
            done = cal.parse_dmy(r.get('tghoanthanh'))
            if cutoff and done and done <= cutoff:
                include = False
                reason = (f'Tạo {m[1]:02d}/{m[0]} và đã hoàn thành {done.strftime("%d/%m %H:%M")} '
                          f'— trước lúc chốt báo cáo tháng trước, nhiều khả năng đã bị loại tay')
                stats['carryover_dropped'] += 1
            else:
                reason = (f'Tạo {m[1]:02d}/{m[0]} nhưng hoàn thành '
                          f'{done.strftime("%d/%m %H:%M") if done else "sau kỳ"} — carryover thật')
                stats['carryover_kept'] += 1

        L, M, N = R.classify(r.get('tag'), r.get('ten'), T, cfg)
        picked.append({
            'id': _id, 'tag': r.get('tag'), 'name': str(r.get('ten') or ''),
            'person': person, 'T': round(T, 2), 'ratings': nums,
            'quality': round(sum(nums), 2),
            'created': str(r.get('ngaytao') or ''), 'done': str(r.get('tghoanthanh') or ''),
            'bu': R.classify_bu(r.get('ten')),
            'L': L, 'M': M, 'N': N,
            'sL': L, 'sM': M, 'sN': N,          # bản đề xuất gốc, để nút "Khôi phục"
            'include': include, 'carryover': is_carry, 'note': reason,
            '_raw': r,
        })

    vers = R.mark_versions(picked)
    for t in picked:
        if t['id'] in vers and t['L'] == 'Video social':
            t['L'] = t['sL'] = 'Video social_Ver'
            t['M'] = t['sM'] = 'Thời gian'
            t['N'] = t['sN'] = t['T']

    stats['picked'] = len(picked)
    stats['included'] = sum(1 for t in picked if t['include'])
    stats['month_label'] = f'{month:02d}/{year}'
    stats['lookback_label'] = cal.label(sorted(prev_months))
    return picked, stats


def score(tasks, cong_chuan, timeoff=0.0, tiers=None):
    """Chấm điểm 1 người từ các task được chọn."""
    tiers = tiers or load_tiers()
    kept = [t for t in tasks if t.get('include')]
    total_n = round(sum(float(t.get('N') or 0) for t in kept), 2)
    denom = (float(cong_chuan) - float(timeoff)) * tiers['hours_per_day']
    time_score = round(total_n / denom * 10, 4) if denom > 0 else 0.0
    qual = round(sum(float(t.get('quality') or 0) for t in kept) / len(kept), 4) if kept else 0.0
    final = round((time_score + qual) / 2, 4)
    tier = next((t for t in tiers['tiers'] if final >= float(t['min'])), tiers['tiers'][-1])
    return {
        'tasks': len(kept), 'total_n': total_n,
        'total_t': round(sum(float(t.get('T') or 0) for t in kept), 2),
        'denominator': round(denom, 2),
        'time_score': round(time_score, 2), 'quality_score': round(qual, 2),
        'final_score': round(final, 2), 'tier': tier['ten'], 'money': tier['tien'],
    }


def bu_summary(tasks, metric='T'):
    """Tổng hợp theo BU (CMA / SC / CEO / ...). metric 'T' = giờ leader phân bổ."""
    kept = [t for t in tasks if t.get('include')]
    agg, per = {}, {}
    for t in kept:
        bu = t.get('bu') or 'Khác'
        v = float(t.get(metric) or 0)
        a = agg.setdefault(bu, {'tasks': 0, 'hours': 0.0})
        a['tasks'] += 1
        a['hours'] = round(a['hours'] + v, 2)
        k = (bu, t['person'])
        p = per.setdefault(k, {'tasks': 0, 'hours': 0.0})
        p['tasks'] += 1
        p['hours'] = round(p['hours'] + v, 2)
    total = round(sum(a['hours'] for a in agg.values()), 2) or 1.0
    rows = []
    for bu in R.BU_ORDER:
        if bu not in agg:
            continue
        rows.append({'bu': bu, 'tasks': agg[bu]['tasks'], 'hours': agg[bu]['hours'],
                     'pct': round(agg[bu]['hours'] / total * 100, 1),
                     'per_person': {p: per[(bu, p)] for (b, p) in per if b == bu}})
    return {'rows': rows, 'total_tasks': len(kept),
            'total_hours': round(sum(a['hours'] for a in agg.values()), 2)}
