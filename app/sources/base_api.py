# -*- coding: utf-8 -*-
"""Nguồn dữ liệu Base.vn qua API — TRẠNG THÁI: BETA, CHƯA KIỂM CHỨNG.

Tài liệu Workflow API của Base.vn nằm trên Postman (render bằng JS) nên không đọc
được tự động, và mình chưa có access_token để thử. Vì vậy module này được viết
theo kiểu "dò rồi map":

  1. `probe()` gọi thử vài endpoint ứng viên, trả về NGUYÊN response để xem shape.
  2. `fetch()` lấy dữ liệu rồi map field theo FIELD_MAP bên dưới.

Khi biết shape thật, chỉ cần sửa ENDPOINTS + FIELD_MAP trong file này, phần còn
lại của app không phải đụng tới.

Lấy token: account.base.vn -> Tích hợp với bên thứ 3 ver 2 -> Thêm application token
(chỉ tài khoản Owner tạo được), chọn app Workflow + scope đọc job.
"""
from __future__ import annotations

import json

import requests

TIMEOUT = 30

# Các endpoint ứng viên, thử lần lượt từ trên xuống.
ENDPOINTS = [
    ('POST', 'https://workflow.base.vn/extapi/v1/job/list'),
    ('POST', 'https://workflow.base.vn/extapi/v1/jobs/list'),
    ('POST', 'https://workflow.base.vn/extapi/v1/job/get'),
]

# field của Base API -> khoá nội bộ của app (giống hệt khoá khi đọc file .xls)
FIELD_MAP = {
    'id': 'id',
    'name': 'ten',
    'since': 'ngaytao',
    'last_update': 'tgcapnhat',
    'complete_time': 'tghoanthanh',
    'content': 'mota',
    'stage_name': 'giaidoan',
    'status': 'trangthai',
}

# các field custom (widget) -> khoá nội bộ; tên thật phụ thuộc cấu hình form Base
CUSTOM_FIELD_MAP = {
    'nguoi_lam_task': 'person',
    'thoi_gian_leader_phan_bo': 'tleader',
    'thoi_gian_hoan_thanh_task': 'tdone',
}

# 5 tiêu chí chấm điểm, theo thứ tự cột V..Z trong báo cáo
RATING_FIELDS = ['danh_gia_deadline', 'danh_gia_thai_do',
                 'tieu_chi_3', 'tieu_chi_4', 'tieu_chi_5']


def probe(access_token: str, extra: dict | None = None) -> list[dict]:
    """Gọi thử các endpoint ứng viên. Trả về kết quả thô để dò shape."""
    out = []
    for method, url in ENDPOINTS:
        payload = {'access_token': access_token}
        payload.update(extra or {})
        item = {'method': method, 'url': url}
        try:
            resp = requests.request(method, url, data=payload, timeout=TIMEOUT)
            item['status'] = resp.status_code
            text = resp.text or ''
            item['raw'] = text[:4000]
            try:
                body = resp.json()
                item['json_keys'] = list(body)[:40] if isinstance(body, dict) else '(list)'
                sample = None
                if isinstance(body, dict):
                    for key in ('jobs', 'data', 'items', 'results'):
                        if isinstance(body.get(key), list) and body[key]:
                            sample = body[key][0]
                            item['list_key'] = key
                            item['count'] = len(body[key])
                            break
                elif isinstance(body, list) and body:
                    sample = body[0]
                    item['count'] = len(body)
                if sample is not None:
                    item['sample_fields'] = list(sample)[:60]
                    item['sample'] = json.dumps(sample, ensure_ascii=False)[:2500]
            except ValueError:
                item['json_keys'] = '(không phải JSON)'
        except Exception as exc:
            item['status'] = None
            item['error'] = f'{type(exc).__name__}: {exc}'
        out.append(item)
    return out


def fetch(access_token: str, endpoint: str, list_key: str = 'jobs',
          tag: str = 'DESIGN MKT', extra: dict | None = None) -> list[dict]:
    """Lấy job rồi map sang shape nội bộ. Gọi được sau khi probe() xác định shape."""
    payload = {'access_token': access_token}
    payload.update(extra or {})
    resp = requests.post(endpoint, data=payload, timeout=TIMEOUT)
    resp.raise_for_status()
    body = resp.json()
    items = body.get(list_key, []) if isinstance(body, dict) else body
    rows = []
    for job in items:
        rec = {v: job.get(k, '') for k, v in FIELD_MAP.items()}
        fields = job.get('form', {}) or job.get('fields', {}) or {}
        if isinstance(fields, list):            # Base hay trả [{id,value}]
            fields = {f.get('id') or f.get('name'): f.get('value') for f in fields}
        for k, v in CUSTOM_FIELD_MAP.items():
            rec[v] = fields.get(k, '')
        rec['ratings'] = [fields.get(k, '') for k in RATING_FIELDS]
        rec['feedback'] = fields.get('feedback', '')
        rec['tag'] = tag
        rec['id'] = str(rec.get('id') or '').strip()
        rec['person'] = str(rec.get('person') or '').strip()
        for k in ('nguoitao', 'giaocho', 'theodoi', 'nhan', 'lydo', 'ngaymuon',
                  'brief', 'phongban', 'kichthuoc', 'mau'):
            rec.setdefault(k, '')
        rows.append(rec)
    return rows
