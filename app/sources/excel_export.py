# -*- coding: utf-8 -*-
"""Đọc file export .xls của Base.vn.

Hai cái bẫy đã gặp thật:
  1. File Base.vn xuất ra bị lỗi OLE -> pandas/openpyxl fail. Phải dùng
     xlrd.open_workbook(..., ignore_workbook_corruption=True).
  2. Số cột KHÔNG cố định giữa các loại order (bản edit-video-lx từng có 30 cột
     do chèn 2 cột học phần ACCA/CFA). Nên map cột THEO TÊN HEADER, tuyệt đối
     không hardcode index.
"""
from __future__ import annotations

import os
import re
import unicodedata

import xlrd

SHEET_NAME = 'Tất cả'

# nhận dạng loại order từ tên file export
TAG_PATTERNS = [
    (r'1-1.*quay-chup',   'QUAYCHUPMKT'),
    (r'1-2.*design',      'DESIGN MKT'),
    (r'1-3.*edit-video',  'VIDEO MKT'),
    (r'1-4.*retouch',     'RETOUCH'),
    (r'2-2.*design',      'DESIGN LX'),
    (r'2-3.*edit-video',  'VIDEO LX'),
]

# cột cần lấy -> các tên header có thể gặp
FIELD_HEADERS = {
    'ngaytao':     ['Ngày tạo'],
    'nguoitao':    ['Người tạo'],
    'giaocho':     ['Giao cho'],
    'ten':         ['Tên nhiệm vụ'],
    'mota':        ['Mô tả nhiệm vụ'],
    'theodoi':     ['Người theo dõi'],
    'giaidoan':    ['Giai đoạn'],
    'trangthai':   ['Trạng thái'],
    'nhan':        ['Nhãn'],
    'lydo':        ['Lý do thất bại'],
    'tghoanthanh': ['Thời gian hoàn thành'],
    'tgcapnhat':   ['Thời gian cập nhật'],
    'ngaymuon':    ['Ngày giờ muốn nhận ấn phẩm', 'Ngày/Giờ muốn nhận ấn phẩm'],
    'brief':       ['Bạn cần Media bám Brief hay tự sáng tạo?'],
    'phongban':    ['Phòng ban giao nhiệm vụ', 'Chọn phòng ban giao nhiệm vụ'],
    'kichthuoc':   ['Kích thước ấn phẩm', 'Kích thước'],
    'mau':         ['Mẫu ấn phẩm (Nếu có)', 'Mẫu ấn phẩm (nếu có)'],
    'person':      ['Người làm task'],
    'tleader':     ['Thời gian leader phân bổ', 'TG Leader phân bổ'],
    'tdone':       ['Thời gian hoàn thành task'],
    'id':          ['ID'],
}

# thứ tự ghi vào Data_Tổng (cột B..R của sheet báo cáo)
RAW_ORDER = ['ngaytao', 'nguoitao', 'giaocho', 'ten', 'mota', 'theodoi', 'giaidoan',
             'trangthai', 'nhan', 'lydo', 'tghoanthanh', 'tgcapnhat', 'ngaymuon',
             'brief', 'phongban', 'kichthuoc', 'mau']


def _norm(s) -> str:
    s = unicodedata.normalize('NFC', str(s))
    return re.sub(r'\s+', ' ', s).strip().lower()


def tag_from_filename(path: str) -> str | None:
    base = os.path.basename(path).lower()
    for pat, tag in TAG_PATTERNS:
        if re.search(pat, base):
            return tag
    return None


def detect_exports(folder: str) -> list[dict]:
    """Quét thư mục, trả về danh sách file export nhận dạng được."""
    out = []
    try:
        names = sorted(os.listdir(folder))
    except OSError:
        return out
    for fn in names:
        if not fn.lower().endswith(('.xls', '.xlsx')):
            continue
        if fn.startswith('~$'):
            continue
        path = os.path.join(folder, fn)
        tag = tag_from_filename(path)
        if tag:
            out.append({'file': fn, 'path': path, 'tag': tag,
                        'size': os.path.getsize(path)})
    return out


def _header_map(sheet) -> dict:
    headers = [str(sheet.cell_value(0, c)) for c in range(sheet.ncols)]
    index = {_norm(h): i for i, h in enumerate(headers)}
    out = {}
    for field, candidates in FIELD_HEADERS.items():
        pos = None
        for cand in candidates:
            if _norm(cand) in index:
                pos = index[_norm(cand)]
                break
        if pos is None:                       # khớp theo tiền tố (header dài bị cắt)
            for cand in candidates:
                key = _norm(cand)[:40]
                for k, v in index.items():
                    if k.startswith(key):
                        pos = v
                        break
                if pos is not None:
                    break
        if pos is None:
            raise ValueError(f'Không tìm thấy cột "{candidates[0]}" trong file export')
        out[field] = pos
    return out


def read_export(path: str, tag: str | None = None) -> list[dict]:
    """Đọc 1 file export -> list dict. 5 cột chấm điểm = 5 cột ngay sau
    'Thời gian hoàn thành task' (tên tiêu chí đổi theo từng loại order)."""
    tag = tag or tag_from_filename(path) or 'KHÁC'
    book = xlrd.open_workbook(path, ignore_workbook_corruption=True)
    sheet = (book.sheet_by_name(SHEET_NAME)
             if SHEET_NAME in book.sheet_names() else book.sheet_by_index(0))
    idx = _header_map(sheet)

    rate0 = idx['tdone'] + 1
    rating_cols = list(range(rate0, rate0 + 5))
    if idx['id'] != rate0 + 6:
        raise ValueError(
            f'Cấu trúc cột lạ ở {os.path.basename(path)}: cột ID ở {idx["id"]}, '
            f'mong đợi {rate0 + 6} (5 cột chấm điểm + 1 cột feedback sau "Thời gian hoàn thành task")')

    rows = []
    for r in range(1, sheet.nrows):
        rec = {f: sheet.cell_value(r, pos) for f, pos in idx.items()}
        rec['tag'] = tag
        rec['ratings'] = [sheet.cell_value(r, c) for c in rating_cols]
        rec['feedback'] = sheet.cell_value(r, rate0 + 5)
        rec['id'] = str(rec['id']).strip()
        rec['person'] = str(rec['person']).strip()
        rec['ten'] = str(rec['ten'])
        rows.append(rec)
    return rows


def read_many(files: list[dict]) -> tuple[list[dict], list[dict]]:
    """files: [{'path','tag'}] -> (rows, report) — report ghi lỗi từng file."""
    rows, report = [], []
    for f in files:
        try:
            got = read_export(f['path'], f.get('tag'))
            rows.extend(got)
            report.append({'file': os.path.basename(f['path']), 'tag': f.get('tag'),
                           'rows': len(got), 'ok': True})
        except Exception as exc:
            report.append({'file': os.path.basename(f['path']), 'tag': f.get('tag'),
                           'rows': 0, 'ok': False, 'error': str(exc)})
    return rows, report
