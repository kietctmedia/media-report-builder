# -*- coding: utf-8 -*-
"""Media Report Builder — web app chạy local (chỉ dùng thư viện chuẩn + openpyxl/xlrd).

Không cần pip install framework: dùng http.server của Python. App chạy local vì
hai lý do bắt buộc: (1) phải đọc file trên ổ E:, (2) phải nhờ Excel tính lại
array formula.
"""
from __future__ import annotations

import io
import json
import mimetypes
import os
import re
import subprocess
import sys
import threading
import traceback
import webbrowser
from email.parser import BytesParser
from email.policy import default as email_default
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs

try:                                   # console Windows mặc định cp1252 -> vỡ tiếng Việt
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')
except Exception:
    pass

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from app.core import builder, calendar_vn as cal, pipeline, rules, template  # noqa: E402
from app.sources import base_api, excel_export  # noqa: E402

WEB_DIR = os.path.join(ROOT, 'app', 'web')
WORKSPACE = os.path.join(ROOT, 'workspace')
os.makedirs(WORKSPACE, exist_ok=True)

STATE: dict = {'tasks': [], 'stats': {}, 'template': None, 'folder': None}
LOCK = threading.Lock()

SLIM = ('id', 'tag', 'name', 'person', 'T', 'quality', 'created', 'done', 'bu',
        'L', 'M', 'N', 'sL', 'sM', 'sN', 'include', 'carryover', 'note')


def slim(tasks):
    return [{k: t[k] for k in SLIM} for t in tasks]


def apply_edits(tasks, edits):
    """edits: {id: {L,M,N,include}} từ client."""
    if not edits:
        return tasks
    for t in tasks:
        e = edits.get(t['id'])
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


# ------------------------------------------------------------------ handlers

def api_scan(body):
    folder = (body.get('folder') or '').strip().strip('"')
    if not folder or not os.path.isdir(folder):
        return {'ok': False, 'error': f'Không thấy thư mục: {folder or "(trống)"}'}
    STATE['folder'] = folder
    return {'ok': True, 'folder': folder,
            'exports': excel_export.detect_exports(folder),
            'reports': template.detect_reports(folder)}


def api_calendar(body):
    y, m = int(body.get('year')), int(body.get('month'))
    return {'ok': True, **cal.suggest_cong_chuan(y, m)}


def api_template(body):
    path = (body.get('path') or '').strip().strip('"')
    if not os.path.isfile(path):
        return {'ok': False, 'error': f'Không thấy file: {path}'}
    try:
        info = template.inspect(path)
    except Exception as exc:
        return {'ok': False, 'error': str(exc)}
    info.pop('ids', None)
    STATE['template'] = path
    return {'ok': True, 'info': info}


def api_analyse(body):
    files = body.get('exports') or []
    tpl = (body.get('template') or '').strip().strip('"')
    year, month = int(body['year']), int(body['month'])
    persons = body.get('persons') or []
    lookback = int(body.get('lookback', 1))
    use_cutoff = bool(body.get('use_cutoff', True))

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
            cutoff = tpl_info['cutoff'] if use_cutoff else None
            tpl_info = {k: v for k, v in tpl_info.items() if k != 'ids'}
        except Exception as exc:
            return {'ok': False, 'error': f'Lỗi đọc file mẫu: {exc}'}

    tasks, stats = pipeline.analyse(rows, year, month, persons, prev_ids, cutoff, lookback)
    with LOCK:
        STATE['tasks'] = tasks
        STATE['stats'] = stats
        STATE['template'] = tpl
    return {'ok': True, 'tasks': slim(tasks), 'stats': stats,
            'read_report': read_report, 'template_info': tpl_info}


def api_score(body):
    with LOCK:
        tasks = apply_edits(STATE['tasks'], body.get('edits') or {})
    cfgs = body.get('persons_cfg') or {}
    tiers = pipeline.load_tiers()
    out = {}
    for person, cfg in cfgs.items():
        sub = [t for t in tasks if t['person'] == person]
        out[person] = pipeline.score(sub, cfg.get('cong_chuan', 22),
                                     cfg.get('timeoff', 0), tiers)
    return {'ok': True, 'scores': out,
            'bu': pipeline.bu_summary(tasks),
            'bu_n': pipeline.bu_summary(tasks, metric='N')}


def api_build(body):
    with LOCK:
        tasks = apply_edits(STATE['tasks'], body.get('edits') or {})
        tpl = (body.get('template') or STATE.get('template') or '').strip().strip('"')
    if not tpl or not os.path.isfile(tpl):
        return {'ok': False, 'error': 'Chưa chọn file báo cáo mẫu của tháng trước.'}
    out_path = (body.get('out_path') or '').strip().strip('"')
    if not out_path:
        return {'ok': False, 'error': 'Chưa đặt đường dẫn file xuất.'}
    if os.path.abspath(out_path) == os.path.abspath(tpl):
        return {'ok': False, 'error': 'File xuất trùng với file mẫu — đổi tên để khỏi ghi đè.'}
    if os.path.exists(out_path) and not body.get('overwrite'):
        return {'ok': False, 'error': 'exists', 'exists': True, 'path': out_path}

    cfgs = body.get('persons_cfg') or {}
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    try:
        layout = builder.build(tasks, tpl, out_path, cfgs)
    except Exception as exc:
        return {'ok': False, 'error': f'Lỗi khi dựng file: {exc}',
                'trace': traceback.format_exc()[-1500:]}

    recalc = builder.recalc(out_path, layout, layout.get('_old_rows'))

    checks = list(recalc.get('checks') or [])
    kept = [t for t in tasks if t['include']]
    checks.insert(0, {'name': 'Số task ghi vào file khớp số task đã chọn',
                      'ok': layout['task_count'] == len(kept),
                      'detail': f"{layout['task_count']} / {len(kept)}"})
    ids = [t['id'] for t in kept]
    checks.insert(1, {'name': 'Không có ID trùng lặp', 'ok': len(ids) == len(set(ids)),
                      'detail': f'{len(ids) - len(set(ids))} trùng'})

    tiers = pipeline.load_tiers()
    preview = {p: pipeline.score([t for t in tasks if t['person'] == p],
                                 c.get('cong_chuan', 22), c.get('timeoff', 0), tiers)
               for p, c in cfgs.items()}
    if recalc.get('available'):
        for person, got in (recalc.get('per_person') or {}).items():
            want = preview.get(person, {}).get('time_score')
            if want is not None:
                checks.append({
                    'name': f'{person}: điểm Excel tính ra khớp bản xem trước',
                    'ok': abs(got['time_score'] - want) < 0.02,
                    'detail': f"Excel {got['time_score']} vs app {want}"})

    return {'ok': True, 'out_path': os.path.abspath(out_path), 'layout_warnings': layout['warnings'],
            'per_person': layout['per_person'], 'recalc': recalc, 'checks': checks,
            'preview': preview, 'bu': pipeline.bu_summary(tasks)}


def api_open(body):
    path = (body.get('path') or '').strip().strip('"')
    if not os.path.exists(path):
        return {'ok': False, 'error': 'Không thấy đường dẫn.'}
    try:
        if body.get('reveal'):
            subprocess.Popen(['explorer', '/select,', os.path.abspath(path)])
        else:
            os.startfile(os.path.abspath(path))  # noqa: S606
        return {'ok': True}
    except Exception as exc:
        return {'ok': False, 'error': str(exc)}


def api_base_probe(body):
    token = (body.get('token') or '').strip()
    if not token:
        return {'ok': False, 'error': 'Chưa nhập access_token.'}
    extra = {}
    if body.get('extra'):
        try:
            extra = json.loads(body['extra'])
        except ValueError:
            return {'ok': False, 'error': 'Tham số thêm không phải JSON hợp lệ.'}
    return {'ok': True, 'results': base_api.probe(token, extra)}


def api_config(body):
    return {'ok': True, 'rate_card': rules.load_rate_card(),
            'tiers': pipeline.load_tiers(),
            'bu_order': rules.BU_ORDER,
            'config_dir': os.path.join(ROOT, 'config')}


ROUTES = {
    '/api/scan': api_scan, '/api/calendar': api_calendar, '/api/template': api_template,
    '/api/analyse': api_analyse, '/api/score': api_score, '/api/build': api_build,
    '/api/open': api_open, '/api/base/probe': api_base_probe, '/api/config': api_config,
}


class Handler(BaseHTTPRequestHandler):
    server_version = 'MediaReportBuilder/1.0'

    def log_message(self, fmt, *args):
        if '/api/' in (args[0] if args else ''):
            sys.stderr.write('  %s\n' % (fmt % args))

    def _send(self, code, payload, ctype='application/json; charset=utf-8'):
        if isinstance(payload, (dict, list)):
            payload = json.dumps(payload, ensure_ascii=False, default=str).encode('utf-8')
        elif isinstance(payload, str):
            payload = payload.encode('utf-8')
        self.send_response(code)
        self.send_header('Content-Type', ctype)
        self.send_header('Content-Length', str(len(payload)))
        self.send_header('Cache-Control', 'no-store')
        self.end_headers()
        self.wfile.write(payload)

    def do_GET(self):
        parsed = urlparse(self.path)
        path = parsed.path
        if path in ('/', '/index.html'):
            return self._file(os.path.join(WEB_DIR, 'index.html'))
        if path.startswith('/static/'):
            rel = path[len('/static/'):]
            safe = os.path.normpath(os.path.join(WEB_DIR, rel))
            if not safe.startswith(WEB_DIR):
                return self._send(403, {'ok': False, 'error': 'forbidden'})
            return self._file(safe)
        if path == '/api/state':
            return self._send(200, {'ok': True, 'tasks': slim(STATE['tasks']),
                                    'stats': STATE['stats'], 'folder': STATE['folder'],
                                    'template': STATE['template']})
        if path == '/api/browse':
            return self._send(200, self._browse(parse_qs(parsed.query)))
        return self._send(404, {'ok': False, 'error': 'not found'})

    def _browse(self, qs):
        start = (qs.get('path') or [''])[0]
        if not start:
            drives = [f'{c}:\\' for c in 'CDEFGH' if os.path.isdir(f'{c}:\\')]
            return {'ok': True, 'path': '', 'parent': None,
                    'dirs': [{'name': d, 'path': d} for d in drives]}
        if not os.path.isdir(start):
            return {'ok': False, 'error': 'Không phải thư mục.'}
        try:
            entries = sorted(e for e in os.listdir(start)
                             if os.path.isdir(os.path.join(start, e)) and not e.startswith('$'))
        except OSError as exc:
            return {'ok': False, 'error': str(exc)}
        parent = os.path.dirname(start.rstrip('\\/')) or None
        if parent == start:
            parent = None
        return {'ok': True, 'path': start, 'parent': parent,
                'dirs': [{'name': e, 'path': os.path.join(start, e)} for e in entries],
                'exports': excel_export.detect_exports(start),
                'reports': template.detect_reports(start)}

    def _file(self, path):
        if not os.path.isfile(path):
            return self._send(404, {'ok': False, 'error': 'not found'})
        ctype = mimetypes.guess_type(path)[0] or 'application/octet-stream'
        if ctype.startswith('text/') or ctype in ('application/javascript',):
            ctype += '; charset=utf-8'
        with open(path, 'rb') as fh:
            self._send(200, fh.read(), ctype)

    def do_POST(self):
        parsed = urlparse(self.path)
        if parsed.path == '/api/upload':
            return self._send(200, self._upload())
        fn = ROUTES.get(parsed.path)
        if not fn:
            return self._send(404, {'ok': False, 'error': 'not found'})
        try:
            length = int(self.headers.get('Content-Length') or 0)
            body = json.loads(self.rfile.read(length) or b'{}')
        except Exception as exc:
            return self._send(400, {'ok': False, 'error': f'Body không hợp lệ: {exc}'})
        try:
            return self._send(200, fn(body))
        except Exception as exc:
            traceback.print_exc()
            return self._send(500, {'ok': False, 'error': f'{type(exc).__name__}: {exc}',
                                    'trace': traceback.format_exc()[-1500:]})

    def _upload(self):
        """Nhận multipart/form-data, lưu vào workspace/ rồi trả đường dẫn thật."""
        ctype = self.headers.get('Content-Type', '')
        if 'multipart/form-data' not in ctype:
            return {'ok': False, 'error': 'Không phải multipart.'}
        length = int(self.headers.get('Content-Length') or 0)
        raw = b'Content-Type: ' + ctype.encode() + b'\r\n\r\n' + self.rfile.read(length)
        msg = BytesParser(policy=email_default).parsebytes(raw)
        saved = []
        for part in msg.iter_parts():
            name = part.get_filename()
            if not name:
                continue
            name = re.sub(r'[\\/:*?"<>|]', '_', os.path.basename(name))
            dest = os.path.join(WORKSPACE, name)
            with open(dest, 'wb') as fh:
                fh.write(part.get_payload(decode=True) or b'')
            saved.append({'file': name, 'path': dest,
                          'tag': excel_export.tag_from_filename(dest),
                          'size': os.path.getsize(dest)})
        STATE['folder'] = WORKSPACE
        return {'ok': True, 'folder': WORKSPACE, 'saved': saved,
                'exports': [s for s in saved if s['tag']],
                'reports': template.detect_reports(WORKSPACE)}


def main():
    port = int(os.environ.get('MRB_PORT', '8765'))
    srv = ThreadingHTTPServer(('127.0.0.1', port), Handler)
    url = f'http://127.0.0.1:{port}/'
    print('=' * 64)
    print('  Media Report Builder')
    print(f'  Đang chạy tại: {url}')
    print('  Đóng cửa sổ này (hoặc Ctrl+C) để tắt app.')
    print('=' * 64)
    if os.environ.get('MRB_NO_BROWSER') != '1':
        threading.Timer(0.8, lambda: webbrowser.open(url)).start()
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        print('\nĐã tắt.')


if __name__ == '__main__':
    main()
