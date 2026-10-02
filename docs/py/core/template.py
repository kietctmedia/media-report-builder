# -*- coding: utf-8 -*-
"""Đọc file báo cáo tháng trước: vừa là TEMPLATE để copy, vừa là nguồn ID để dedup.

Không hardcode vị trí sheet/cột — dò theo tên sheet và tên header, để báo cáo có
đổi layout chút đỉnh thì app vẫn chạy.
"""
from __future__ import annotations

import os
import re

import openpyxl

from . import calendar_vn as cal

DATA_SHEET = 'Data_Tổng'
MASTER_SHEET = 'Master'
SUFFIX = '_Rút gọn'

# header hàng 1 của Data_Tổng -> khoá nội bộ
DATA_HEADERS = {
    'ngày tạo': 'ngaytao', 'tên nhiệm vụ': 'ten', 'người làm task': 'person',
    'thời gian leader phân bổ': 'tleader', 'id': 'id',
    'thời gian hoàn thành': 'tghoanthanh',
}


def _norm(s) -> str:
    return re.sub(r'\s+', ' ', str(s or '')).strip().lower()


def detect_reports(folder: str) -> list[dict]:
    """Tìm các file .xlsx trông giống báo cáo (có sheet Data_Tổng)."""
    out = []
    try:
        names = sorted(os.listdir(folder))
    except OSError:
        return out
    for fn in names:
        if not fn.lower().endswith('.xlsx') or fn.startswith('~$'):
            continue
        path = os.path.join(folder, fn)
        try:
            wb = openpyxl.load_workbook(path, read_only=True)
            ok = DATA_SHEET in wb.sheetnames
            sheets = list(wb.sheetnames)
            wb.close()
        except Exception:
            continue
        if ok:
            out.append({'file': fn, 'path': path, 'sheets': sheets,
                        'mtime': os.path.getmtime(path)})
    out.sort(key=lambda x: x['mtime'], reverse=True)
    return out


def inspect(path: str) -> dict:
    """Trả về cấu trúc + danh sách ID + mốc thời gian của báo cáo tháng trước."""
    wbv = openpyxl.load_workbook(path, data_only=True)
    if DATA_SHEET not in wbv.sheetnames:
        raise ValueError(f'File không có sheet "{DATA_SHEET}" — '
                         f'đây có vẻ không phải file báo cáo mẫu.')
    ws = wbv[DATA_SHEET]

    cols = {}
    for c in range(1, ws.max_column + 1):
        key = DATA_HEADERS.get(_norm(ws.cell(1, c).value))
        if key and key not in cols:
            cols[key] = c
    for need in ('person', 'tleader', 'id', 'ten'):
        if need not in cols:
            raise ValueError(f'Data_Tổng thiếu cột "{need}" — không dò được header.')

    last = 1
    for r in range(2, ws.max_row + 1):
        if ws.cell(r, 1).value not in (None, ''):
            last = r

    ids, people, cutoff = [], {}, None
    for r in range(2, last + 1):
        _id = str(ws.cell(r, cols['id']).value or '').strip()
        if not _id:
            continue
        ids.append(_id)
        p = str(ws.cell(r, cols['person']).value or '').strip()
        if p:
            people[p] = people.get(p, 0) + 1
        if 'tghoanthanh' in cols:
            d = cal.parse_dmy(ws.cell(r, cols['tghoanthanh']).value)
            if d and (cutoff is None or d > cutoff):
                cutoff = d

    persons = [s[:-len(SUFFIX)] for s in wbv.sheetnames if s.endswith(SUFFIX)]
    wbv.close()

    return {
        'path': path,
        'file': os.path.basename(path),
        'data_last_row': last,
        'task_count': len(ids),
        'ids': ids,
        'persons': persons or sorted(people, key=lambda k: -people[k]),
        'person_counts': people,
        'cutoff': cutoff.isoformat() if cutoff else None,
        'cutoff_label': cutoff.strftime('%d/%m/%Y %H:%M') if cutoff else None,
        'has_master': MASTER_SHEET in openpyxl.load_workbook(path, read_only=True).sheetnames,
    }


def read_lmn(path: str) -> dict:
    """Đọc L/M/N theo ID từ các sheet Rút gọn (để giữ chỉnh tay khi làm bản Ver)."""
    wbv = openpyxl.load_workbook(path, data_only=True)
    wbf = openpyxl.load_workbook(path)
    out = {}
    for sheet in wbv.sheetnames:
        if not sheet.endswith(SUFFIX):
            continue
        wsv, wsf = wbv[sheet], wbf[sheet]
        for r in range(3, wsf.max_row + 1):
            _id = str(wsv.cell(r, 11).value or '').strip()
            if not _id:
                continue
            out[_id] = {
                'L': wsf.cell(r, 12).value,
                'M': wsf.cell(r, 13).value,
                'N': wsf.cell(r, 14).value,
            }
    wbv.close()
    wbf.close()
    return out
