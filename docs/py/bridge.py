# -*- coding: utf-8 -*-
"""Cầu nối cho bản chạy thẳng trong trình duyệt (Pyodide).

Nhận đúng những "đường dẫn API" mà app.js vẫn gọi, nhưng chạy tại chỗ thay vì
gửi lên máy chủ. Dùng chung y hệt app/core và app/sources với bản chạy trên máy,
nên logic nghiệp vụ chỉ có MỘT bản duy nhất.
"""
from __future__ import annotations

import json
import os
import traceback

from app.core import builder, calendar_vn as cal, pipeline, rules, template
from app.sources import excel_export

WORK = '/work'
STATE: dict = {'tasks': [], 'stats': {}, 'template': None, 'folder': WORK}

SLIM = ('id', 'tag', 'name', 'person', 'T', 'quality', 'created', 'done', 'bu',
        'L', 'M', 'N', 'sL', 'sM', 'sN', 'include', 'carryover', 'note')


def _slim(tasks):
    return [{k: t[k] for k in SLIM} for t in tasks]


def _apply_edits(tasks, edits):
    for t in tasks:
        e = (edits or {}).get(t['id'])
        if not e:
            continue
        if 'L' in e:
            t['L'] = e['L']
        if 'M' in e:
            t['M'] = e['M']
        if 'N' in e:
            try:
                t['N'] = round(float(e['N']), 2)
            except (TypeError, ValueError):
                pass
        if 'include' in e:
            t['include'] = bool(e['include'])
    return tasks


# ----------------------------------------------------------------- các route

def _config(body):
    return {'ok': True, 'cloud': True, 'static': True, 'excel': False,
            'rate_card': rules.load_rate_card(), 'tiers': pipeline.load_tiers(),
            'bu_order': rules.BU_ORDER}


def _state(body):
    return {'ok': True, 'tasks': _slim(STATE['tasks']), 'stats': STATE['stats'],
            'folder': STATE['folder'], 'template': STATE['template']}


def _scan(body):
    return {'ok': False, 'error': 'Bản chạy trong trình duyệt không quét ổ đĩa — hãy chọn file.'}


def _calendar(body):
    return {'ok': True, **cal.suggest_cong_chuan(int(body['year']), int(body['month']))}


def _files(body):
    """Liệt kê lại những file đã nạp vào bộ nhớ ảo."""
    return {'ok': True, 'folder': WORK,
            'exports': excel_export.detect_exports(WORK),
            'reports': template.detect_reports(WORK)}


def _template(body):
    path = body.get('path') or ''
    if not os.path.isfile(path):
        return {'ok': False, 'error': f'Không thấy file: {os.path.basename(path)}'}
    try:
        info = template.inspect(path)
    except Exception as exc:
        return {'ok': False, 'error': str(exc)}
    info.pop('ids', None)
    STATE['template'] = path
    return {'ok': True, 'info': info}


def _analyse(body):
    files = body.get('exports') or []
    tpl = body.get('template') or ''
    persons = body.get('persons') or []
    if not files:
        return {'ok': False, 'error': 'Chưa chọn file dữ liệu export nào.'}
    if not persons:
        return {'ok': False, 'error': 'Chưa chọn người để đánh giá.'}

    rows, read_report = excel_export.read_many(files)
    if not rows:
        return {'ok': False, 'error': 'Không đọc được dòng nào từ các file export.',
                'read_report': read_report}

    prev_ids, cutoff, tpl_info = [], None, None
    if tpl and os.path.isfile(tpl):
        try:
            tpl_info = template.inspect(tpl)
            prev_ids = tpl_info['ids']
            cutoff = tpl_info['cutoff'] if body.get('use_cutoff', True) else None
            tpl_info = {k: v for k, v in tpl_info.items() if k != 'ids'}
        except Exception as exc:
            return {'ok': False, 'error': f'Lỗi đọc file mẫu: {exc}'}

    tasks, stats = pipeline.analyse(rows, int(body['year']), int(body['month']),
                                    persons, prev_ids, cutoff,
                                    int(body.get('lookback', 1)))
    STATE.update(tasks=tasks, stats=stats, template=tpl)
    return {'ok': True, 'tasks': _slim(tasks), 'stats': stats,
            'read_report': read_report, 'template_info': tpl_info}


def _score(body):
    tasks = _apply_edits(STATE['tasks'], body.get('edits'))
    tiers = pipeline.load_tiers()
    scores = {p: pipeline.score([t for t in tasks if t['person'] == p],
                                c.get('cong_chuan', 22), c.get('timeoff', 0), tiers)
              for p, c in (body.get('persons_cfg') or {}).items()}
    return {'ok': True, 'scores': scores,
            'bu': pipeline.bu_summary(tasks),
            'bu_n': pipeline.bu_summary(tasks, metric='N')}


def _build(body):
    tasks = _apply_edits(STATE['tasks'], body.get('edits'))
    tpl = body.get('template') or STATE.get('template') or ''
    if not tpl or not os.path.isfile(tpl):
        return {'ok': False, 'error': 'Chưa chọn file báo cáo mẫu của tháng trước.'}

    name = os.path.basename(body.get('out_path') or 'bao-cao.xlsx') or 'bao-cao.xlsx'
    if not name.lower().endswith('.xlsx'):
        name += '.xlsx'
    out_path = os.path.join(WORK, '__out__', name)
    os.makedirs(os.path.dirname(out_path), exist_ok=True)

    cfgs = body.get('persons_cfg') or {}
    try:
        layout = builder.build(tasks, tpl, out_path, cfgs)
    except Exception as exc:
        return {'ok': False, 'error': f'Lỗi khi dựng file: {exc}',
                'trace': traceback.format_exc()[-1200:]}

    kept = [t for t in tasks if t['include']]
    ids = [t['id'] for t in kept]
    checks = [
        {'name': 'Số task ghi vào file khớp số task đã chọn',
         'ok': layout['task_count'] == len(kept),
         'detail': f"{layout['task_count']} / {len(kept)}"},
        {'name': 'Không có ID trùng lặp', 'ok': len(ids) == len(set(ids)),
         'detail': f'{len(ids) - len(set(ids))} trùng'},
    ]
    for person, info in layout['per_person'].items():
        want = layout['_sum_t'].get(person, 0.0)
        got = sum(float(t['T']) for t in kept if t['person'] == person)
        checks.append({'name': f'{person}: tổng giờ phân bổ ghi vào Data_Tổng khớp',
                       'ok': abs(got - want) < 0.01, 'detail': f'{got} vs {want}'})
        checks.append({'name': f'{person}: số dòng khớp số task',
                       'ok': info['count'] == len([t for t in kept if t['person'] == person]),
                       'detail': str(info['count'])})

    tiers = pipeline.load_tiers()
    preview = {p: pipeline.score([t for t in tasks if t['person'] == p],
                                 c.get('cong_chuan', 22), c.get('timeoff', 0), tiers)
               for p, c in cfgs.items()}

    return {'ok': True, 'out_path': name, 'file_path': out_path,
            'layout_warnings': layout['warnings'], 'per_person': layout['per_person'],
            'recalc': {'available': False,
                       'reason': 'Trình duyệt không có Excel — số sẽ hiện khi bạn mở file'},
            'checks': checks, 'preview': preview, 'bu': pipeline.bu_summary(tasks)}


ROUTES = {
    '/api/config': _config, '/api/state': _state, '/api/scan': _scan,
    '/api/calendar': _calendar, '/api/template': _template, '/api/analyse': _analyse,
    '/api/score': _score, '/api/build': _build, '/api/files': _files,
}


def call(path: str, body_json: str) -> str:
    """Điểm vào duy nhất mà JavaScript gọi."""
    try:
        body = json.loads(body_json or '{}')
    except ValueError as exc:
        return json.dumps({'ok': False, 'error': f'Body không hợp lệ: {exc}'})
    fn = ROUTES.get(path)
    if fn is None:
        return json.dumps({'ok': False, 'error': f'Không có route {path}'})
    try:
        return json.dumps(fn(body), ensure_ascii=False, default=str)
    except Exception as exc:
        return json.dumps({'ok': False, 'error': f'{type(exc).__name__}: {exc}',
                           'trace': traceback.format_exc()[-1200:]}, ensure_ascii=False)
