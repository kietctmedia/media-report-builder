# -*- coding: utf-8 -*-
"""Chạy lại nguyên pipeline trên dữ liệu tháng 9/2026 có thật và so với báo cáo
đã làm tay, để chắc app cho ra cùng kết quả.

    python selftest.py            # chỉ kiểm tra logic, không ghi file
    python selftest.py --build    # dựng thêm file .xlsx ra workspace/
"""
from __future__ import annotations

import os
import sys

try:                                   # console Windows mặc định cp1252 -> vỡ tiếng Việt
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')
except Exception:
    pass

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from app.core import builder, pipeline, template  # noqa: E402
from app.sources import excel_export  # noqa: E402

FOLDER = r'E:\Báo cáo hiệu suất\2026\0926'
TEMPLATE = os.path.join(FOLDER, 'MEDIA_CMA_Báo cáo hiệu suất team T8_2026.xlsx')
EXPECT = {'tasks': 132, 'Kiệt': {'n': 54, 'T': 252.0}, 'Giang': {'n': 78, 'T': 207.1}}
CFG = {'Kiệt': {'cong_chuan': 20, 'timeoff': 0, 'ot': 0},
       'Giang': {'cong_chuan': 20, 'timeoff': 0, 'ot': 0}}

fails = []


def check(name, ok, detail=''):
    print(f'  [{"OK " if ok else "SAI"}] {name}' + (f'  — {detail}' if detail else ''))
    if not ok:
        fails.append(name)


def main():
    print('\n== 1. Nhận diện file trong thư mục ==')
    exports = excel_export.detect_exports(FOLDER)
    for e in exports:
        print(f'   {e["tag"]:12} {e["file"]}')
    check('Nhận ra 5 file export', len(exports) == 5, f'{len(exports)} file')

    reports = template.detect_reports(FOLDER)
    check('Nhận ra file báo cáo mẫu', any(r['path'] == TEMPLATE for r in reports),
          f'{len(reports)} file .xlsx có Data_Tổng')

    print('\n== 2. Đọc template tháng trước ==')
    info = template.inspect(TEMPLATE)
    check('Đọc đúng số task của T8', info['task_count'] == 153, f'{info["task_count"]}')
    check('Dò ra 2 nhân sự từ tên sheet', set(info['persons']) == {'Kiệt', 'Giang'},
          str(info['persons']))
    print(f'   mốc chốt kỳ trước: {info["cutoff_label"]}')

    print('\n== 3. Đọc dữ liệu export ==')
    rows, report = excel_export.read_many(
        [{'path': e['path'], 'tag': e['tag']} for e in exports])
    for r in report:
        if not r['ok']:
            print(f'   LỖI {r["file"]}: {r["error"]}')
    check('Mọi file đọc được', all(r['ok'] for r in report))
    check('Tổng số dòng thô', len(rows) == 4655, f'{len(rows)} dòng')

    print('\n== 4. Lọc + dedup ==')
    tasks, stats = pipeline.analyse(rows, 2026, 9, ['Kiệt', 'Giang'],
                                    info['ids'], info['cutoff'], lookback=1)
    kept = [t for t in tasks if t['include']]
    check('Số task sau lọc', len(kept) == EXPECT['tasks'],
          f'{len(kept)} (mong đợi {EXPECT["tasks"]})')
    for p, exp in (('Kiệt', EXPECT['Kiệt']), ('Giang', EXPECT['Giang'])):
        sub = [t for t in kept if t['person'] == p]
        tt = round(sum(t['T'] for t in sub), 2)
        check(f'{p}: số task', len(sub) == exp['n'], f'{len(sub)} vs {exp["n"]}')
        check(f'{p}: tổng giờ phân bổ', abs(tt - exp['T']) < 0.01, f'{tt} vs {exp["T"]}')
    print(f'   carryover giữ {stats["carryover_kept"]} / tự bỏ {stats["carryover_dropped"]}')
    print(f'   trùng với báo cáo trước: {stats["duplicated"]}')

    print('\n== 5. Chấm điểm (đối chiếu bản làm tay) ==')
    for p, want_t, want_q in (('Kiệt', 15.44, 9.95), ('Giang', 12.94, 10.0)):
        sc = pipeline.score([t for t in kept if t['person'] == p],
                            CFG[p]['cong_chuan'], CFG[p]['timeoff'])
        check(f'{p}: điểm thời gian', abs(sc['time_score'] - want_t) < 0.02,
              f'{sc["time_score"]} vs {want_t}')
        check(f'{p}: điểm chất lượng', abs(sc['quality_score'] - want_q) < 0.02,
              f'{sc["quality_score"]} vs {want_q}')
        print(f'   -> {p}: {sc["final_score"]} · {sc["tier"]} · {sc["money"]:,}đ')

    print('\n== 6. Phân tích BU ==')
    bu = pipeline.bu_summary(kept)
    for r in bu['rows']:
        print(f'   {r["bu"]:10} {r["tasks"]:4} task  {r["hours"]:7.1f} h  {r["pct"]:5.1f}%')
    cma = next((r for r in bu['rows'] if r['bu'] == 'CMA'), None)
    check('BU CMA khớp bản làm tay', cma and cma['tasks'] == 68 and abs(cma['hours'] - 214.1) < 0.05,
          f'{cma["tasks"]} task / {cma["hours"]}h' if cma else 'không có')

    if '--build' in sys.argv:
        print('\n== 7. Dựng file .xlsx ==')
        out = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                           'workspace', 'selftest_T9.xlsx')
        os.makedirs(os.path.dirname(out), exist_ok=True)
        layout = builder.build(kept, TEMPLATE, out, CFG)
        check('Ghi đủ task', layout['task_count'] == len(kept),
              f'{layout["task_count"]}')
        rec = builder.recalc(out, layout, layout.get('_old_rows'))
        if rec.get('available'):
            for p, v in rec['per_person'].items():
                print(f'   Excel tính {p}: thời gian {v["time_score"]} · chất lượng {v["quality_score"]}')
            bad = [c for c in rec['checks'] if not c['ok']]
            check('Kiểm tra trong file đều đạt', not bad,
                  '; '.join(f'{c["name"]} ({c["detail"]})' for c in bad))
            for p, want in (('Kiệt', 15.44), ('Giang', 12.94)):
                got = rec['per_person'][p]['time_score']
                check(f'{p}: Excel khớp bản làm tay', abs(got - want) < 0.02, f'{got} vs {want}')
        else:
            print(f'   (bỏ qua recalc: {rec.get("reason")})')
        print(f'   file: {out}')

    print('\n' + '=' * 62)
    if fails:
        print(f'  {len(fails)} KIỂM TRA KHÔNG ĐẠT:')
        for f in fails:
            print('   -', f)
        sys.exit(1)
    print('  TẤT CẢ KIỂM TRA ĐỀU ĐẠT')
    print('=' * 62)


if __name__ == '__main__':
    main()
