# -*- coding: utf-8 -*-
"""Sinh thư mục docs/ — bản tĩnh chạy thẳng trong trình duyệt, để GitHub Pages host.

Dùng CHUNG app/core và app/sources với bản chạy trên máy: script này chỉ copy,
không có bản sao logic nghiệp vụ nào cả. Chạy lại mỗi khi sửa core hoặc giao diện:

    python build_pages.py
"""
from __future__ import annotations

import os
import shutil
import sys

try:
    sys.stdout.reconfigure(encoding='utf-8')
except Exception:
    pass

ROOT = os.path.dirname(os.path.abspath(__file__))
WEB = os.path.join(ROOT, 'app', 'web')
DOCS = os.path.join(ROOT, 'docs')

PY_MODULES = [
    ('app/core/calendar_vn.py', 'py/core/calendar_vn.py'),
    ('app/core/rules.py', 'py/core/rules.py'),
    ('app/core/pipeline.py', 'py/core/pipeline.py'),
    ('app/core/template.py', 'py/core/template.py'),
    ('app/core/builder.py', 'py/core/builder.py'),
    ('app/sources/excel_export.py', 'py/sources/excel_export.py'),
    ('app/web/bridge.py', 'py/bridge.py'),
]
CONFIGS = ['rate_card.json', 'kpi_tiers.json', 'holidays.json']

PYODIDE = '<script src="https://cdn.jsdelivr.net/pyodide/v0.28.0/full/pyodide.js"></script>'


def build():
    if os.path.isdir(DOCS):
        shutil.rmtree(DOCS)
    for sub in ('py/core', 'py/sources', 'config'):
        os.makedirs(os.path.join(DOCS, sub), exist_ok=True)

    # giao diện
    shutil.copyfile(os.path.join(WEB, 'style.css'), os.path.join(DOCS, 'style.css'))
    shutil.copyfile(os.path.join(WEB, 'app.js'), os.path.join(DOCS, 'app.js'))
    shutil.copyfile(os.path.join(WEB, 'transport-pyodide.js'),
                    os.path.join(DOCS, 'transport-pyodide.js'))

    html = open(os.path.join(WEB, 'index.html'), encoding='utf-8').read()
    html = html.replace('/static/style.css', 'style.css')
    html = html.replace('/static/app.js', 'app.js')
    html = html.replace('<!-- TRANSPORT --><script src="/static/transport-server.js"></script>',
                        PYODIDE + '\n<script src="transport-pyodide.js"></script>')
    assert 'transport-pyodide.js' in html, 'Không thay được dòng transport trong index.html'
    html = html.replace('<title>Media Report Builder</title>',
                        '<title>Media Report Builder</title>\n'
                        '<meta name="description" content="Dựng báo cáo hiệu suất team Media '
                        'từ file export Base.vn — chạy hoàn toàn trong trình duyệt.">')
    open(os.path.join(DOCS, 'index.html'), 'w', encoding='utf-8').write(html)

    # module python + config
    for src, dst in PY_MODULES:
        shutil.copyfile(os.path.join(ROOT, src), os.path.join(DOCS, dst))
    for c in CONFIGS:
        shutil.copyfile(os.path.join(ROOT, 'config', c), os.path.join(DOCS, 'config', c))

    # GitHub Pages: khỏi chạy Jekyll (nếu không, thư mục/underscore bị bỏ qua)
    open(os.path.join(DOCS, '.nojekyll'), 'w').close()

    n = sum(len(f) for _, _, f in os.walk(DOCS))
    size = sum(os.path.getsize(os.path.join(r, f))
               for r, _, fs in os.walk(DOCS) for f in fs)
    print(f'Đã sinh docs/ — {n} file, {size / 1024:.0f} KB')
    for r, _, fs in sorted(os.walk(DOCS)):
        rel = os.path.relpath(r, DOCS).replace('\\', '/')
        for f in sorted(fs):
            print(f'   {"" if rel == "." else rel + "/"}{f}')


if __name__ == '__main__':
    build()
