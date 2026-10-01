# -*- coding: utf-8 -*-
"""Phân loại task -> (L = Phân loại, M = Standard, N = Total giờ).

Quy tắc N chốt từ bản T8/2026 mà leader tự cân tay:
    N = min(T, định mức)
tức là giờ leader phân bổ bị chặn trên bởi định mức rate card. Nhãn nào không có
định mức (M = 'Thời gian') thì N = T.

Toàn bộ số định mức nằm ở config/rate_card.json.
"""
from __future__ import annotations

import json
import os
import re

_CFG_PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(
    os.path.abspath(__file__)))), 'config', 'rate_card.json')

_DEFAULTS = {
    'anh_post_gio_moi_anh': 2.0,
    'bo_anh_social_gio_moi_anh': 0.5,
    'resize_gio_moi_anh': 0.1667,
    'kv_gio': 4.0,
    'landingpage_gio_moi_section': 0.6667,
    'quay_lop_hoc_gio': 3.0,
    'chup_lop_hoc_gio': 2.0,
    'video_social_gio': 6.0,
    'video_animation_gio': 6.0,
    'video_sx_gio_moi_buoi': 3.0,
}


def load_rate_card() -> dict:
    cfg = dict(_DEFAULTS)
    try:
        with open(_CFG_PATH, encoding='utf-8') as fh:
            raw = json.load(fh)
        for k, v in raw.items():
            if k.startswith('_'):
                continue
            if isinstance(v, (int, float)):
                cfg[k] = float(v)
        cfg['nguong_anh_post'] = float(
            raw.get('_nguong_anh_post', {}).get('gio_moi_anh', 2))
        cfg['ap_dung_dinh_muc'] = bool(
            raw.get('_ap_dung_dinh_muc', {}).get('bat', True))
    except Exception:
        cfg['nguong_anh_post'] = 2.0
        cfg['ap_dung_dinh_muc'] = True
    return cfg


def _n(count: int | None, per_unit: float, T: float, cfg: dict) -> float:
    """N = min(T, định mức) — làm tròn 2 chữ số để tránh rác float."""
    if not cfg.get('ap_dung_dinh_muc', True):
        return round(T, 2)
    quota = round((count or 1) * per_unit, 2)
    return round(min(T, quota), 2)


def _count(name: str, unit: str) -> int | None:
    m = re.search(r'(\d+)\s*' + unit, str(name).lower())
    return int(m.group(1)) if m else None


# ---------------------------------------------------------------- tag handlers

def _design(name: str, n: str, T: float, cfg: dict):
    if 'resize' in n:
        c = _count(name, 'ảnh') or 1
        return f'Ảnh social_Resize_{c} ảnh', '10p/ảnh', round(T, 2)
    if 'design system' in n:
        return 'Design System', 'Thời gian', round(T, 2)
    if re.search(r'(?:^|[\s_])s[ửu]a|chỉnh sửa|tách nền|tách element|lọc,', n):
        return 'Ảnh social_Sửa', 'Thời gian', round(T, 2)
    if re.search(r'(?:^|[\s_])kv(?:[\s_]|$)', n) or 'key visual' in n:
        return 'KV', '4h', _n(1, cfg['kv_gio'], T, cfg)
    sec = _count(name, 'section')
    if sec:
        return (f'Landingpage_{sec} section', '40p/section',
                _n(sec, cfg['landingpage_gio_moi_section'], T, cfg))
    c = _count(name, 'ảnh')
    if c is None:
        # tên task không ghi số lượng (moodboard, ebook, "design lại định vị"...)
        # -> không suy ra được định mức, tin giờ leader đã cấp
        return 'Ảnh post', '2h/ảnh', round(T, 2)
    if c == 1:
        return 'Ảnh post', '2h/ảnh', _n(1, cfg['anh_post_gio_moi_anh'], T, cfg)
    if T / c >= cfg['nguong_anh_post']:
        return 'Ảnh post', '2h/ảnh', _n(c, cfg['anh_post_gio_moi_anh'], T, cfg)
    return (f'Bộ ảnh social_{c} ảnh', '30p/ảnh',
            _n(c, cfg['bo_anh_social_gio_moi_anh'], T, cfg))


def _quaychup(name: str, n: str, T: float, cfg: dict):
    if 'retouch' in n:
        return 'Retouch', 'Thời gian', round(T, 2)
    if re.search(r'(?:^|[\s_])sc\d*(?:[\s_]|$)', n) and 'kg' in n and T >= 6:
        return 'Chụp lớp SC + quay highlight + quay phỏng vấn', 'Thời gian', round(T, 2)
    if re.search(r'lớp|trial|khai giảng|masterclass', n):
        if 'quay' in n:
            return 'Quay lớp học', '3h', _n(1, cfg['quay_lop_hoc_gio'], T, cfg)
        return 'Chụp lớp học', 'Thời gian', round(T, 2)
    if 'chụp' in n:
        return 'Chụp sự kiện', 'Thời gian', round(T, 2)
    if 'quay' in n:
        return 'Quay sự kiện', 'Thời gian', round(T, 2)
    return 'Chụp sự kiện', 'Thời gian', round(T, 2)


def _video_mkt(name: str, n: str, T: float, cfg: dict):
    if 'retouch' in n:
        return 'Retouch', 'Thời gian', round(T, 2)
    if re.search(r'claude|premiere|setup|skill|agent', n):
        return 'Video social_Setup AI', 'Thời gian', round(T, 2)
    if re.search(r'(?:^|[\s_])s[ửu]a', n):
        return 'Video social_Sửa', 'Thời gian', round(T, 2)
    if 'xào lại' in n:
        return 'Video social_Ver', 'Thời gian', round(T, 2)
    if 'animation' in n:
        return 'Video animation', '6h', _n(1, cfg['video_animation_gio'], T, cfg)
    return 'Video social', '6h', _n(1, cfg['video_social_gio'], T, cfg)


def _video_lx(name: str, n: str, T: float, cfg: dict):
    if 'convert' in n:
        return 'Convert record', 'Thời gian', round(T, 2)
    per = cfg['video_sx_gio_moi_buoi']
    if T < per:
        return 'Video SX_nhỏ lẻ', 'Thời gian', round(T, 2)
    buoi = max(1, int(round(T / per)))
    return f'Video SX_{buoi} buổi', '3h', round(T, 2)


_HANDLERS = {
    'DESIGN MKT': _design,
    'DESIGN LX': _design,
    'QUAYCHUPMKT': _quaychup,
    'VIDEO MKT': _video_mkt,
    'VIDEO LX': _video_lx,
    'RETOUCH': lambda name, n, T, cfg: ('Retouch', 'Thời gian', round(T, 2)),
}


def classify(tag: str, name: str, T: float, cfg: dict | None = None):
    cfg = cfg or load_rate_card()
    try:
        T = float(T)
    except (TypeError, ValueError):
        T = 0.0
    fn = _HANDLERS.get(tag)
    if fn is None:
        return 'Khác', 'Thời gian', round(T, 2)
    return fn(name, str(name).lower(), T, cfg)


# ----------------------------------------------------------- hậu xử lý: bản ver

_VER_RE = re.compile(r'^(.*?)[\s_]+(?:video[\s_]+)?(\d+)$', re.I)


def mark_versions(rows: list[dict]) -> set:
    """Đánh dấu các bản ver đánh số của cùng một video gốc ('... A4 2', '... video 1').

    rows: [{'id','tag','name'}] -> set(id)
    """
    names = {str(r['name']).strip().lower()
             for r in rows if r.get('tag') == 'VIDEO MKT'}
    out = set()
    for r in rows:
        if r.get('tag') != 'VIDEO MKT':
            continue
        m = _VER_RE.match(str(r['name']).strip())
        if m and m.group(1).strip().lower() in names:
            out.add(r['id'])
    return out


# --------------------------------------------------------------- phân loại BU

_BU_RULES = [
    ('SC',        lambda n: bool(re.search(r'(?:^|[\s_])sc\d*(?:[\s_]|$)', n))
                            or 'shortcourse' in n or 'short course' in n),
    # CGMA / B2B xét TRƯỚC CEO: task "0926_B2B_Book lịch quay Event CEO Forum"
    # thuộc B2B chứ không phải lớp CEO.
    ('CGMA',      lambda n: bool(re.search(r'(?:^|[\s_])cgma(?:[\s_]|$)', n))),
    ('B2B',       lambda n: bool(re.search(r'(?:^|[\s_])b2b(?:[\s_]|$)', n))),
    ('CEO',       lambda n: bool(re.search(r'(?:^|[\s_])ceo(?:[\s_]|$)', n))
                            or bool(re.search(r'(?:^|[\s_])cmc\s*\d*(?:[\s_]|$)', n))),
    ('FASIS',     lambda n: 'fasis' in n),
    ('HO/Nội bộ', lambda n: 'office tour' in n
                            or bool(re.search(r'(?:^|[\s_])ho(?:[\s_]|$)', n))),
    ('CMA',       lambda n: bool(re.search(r'(?:^|[\s_])cma(?:[\s_]|$)', n))),
]

BU_ORDER = ['CMA', 'SC', 'CEO', 'CGMA', 'B2B', 'FASIS', 'HO/Nội bộ', 'Khác']


def classify_bu(name: str) -> str:
    n = str(name).lower()
    for bu, test in _BU_RULES:
        if test(n):
            return bu
    return 'Khác'
