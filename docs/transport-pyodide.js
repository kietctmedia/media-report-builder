/* Transport cho bản tĩnh (GitHub Pages): chạy thẳng Python trong trình duyệt
   bằng Pyodide. Không có máy chủ, file của bạn KHÔNG rời khỏi máy. */
'use strict';

const PYODIDE_URL = 'https://cdn.jsdelivr.net/pyodide/v0.28.0/full/';
const PY_MODULES = [
  'core/calendar_vn.py', 'core/rules.py', 'core/pipeline.py',
  'core/template.py', 'core/builder.py', 'sources/excel_export.py',
];
const CONFIGS = ['rate_card.json', 'kpi_tiers.json', 'holidays.json'];
const WORK = '/work';

let _py = null;
let _booting = null;

function boot(onStep) {
  if (_py) return Promise.resolve(_py);
  if (_booting) return _booting;
  _booting = (async () => {
    onStep && onStep('Đang tải Python cho trình duyệt…');
    const py = await loadPyodide({ indexURL: PYODIDE_URL });
    onStep && onStep('Đang cài openpyxl và xlrd…');
    await py.loadPackage('micropip');
    await py.pyimport('micropip').install(['openpyxl', 'xlrd==2.0.1']);

    onStep && onStep('Đang nạp mã xử lý…');
    py.FS.mkdirTree(WORK);
    py.FS.mkdirTree('/app/core');
    py.FS.mkdirTree('/app/sources');
    py.FS.mkdirTree('/config');
    py.FS.writeFile('/app/__init__.py', '');
    py.FS.writeFile('/app/core/__init__.py', '');
    py.FS.writeFile('/app/sources/__init__.py', '');
    await Promise.all(PY_MODULES.map(async m => {
      py.FS.writeFile('/app/' + m, await (await fetch('py/' + m)).text());
    }));
    await Promise.all(CONFIGS.map(async c => {
      py.FS.writeFile('/config/' + c, await (await fetch('config/' + c)).text());
    }));
    py.FS.writeFile('/bridge.py', await (await fetch('py/bridge.py')).text());

    py.runPython("import sys; sys.path.insert(0, '/')");
    py.runPython('import bridge');
    _py = py;
    return py;
  })();
  return _booting;
}

function setMsg(text) {
  const el = document.getElementById('overlayMsg');
  const ov = document.getElementById('overlay');
  if (ov) ov.classList.remove('hidden');
  if (el) el.textContent = text;
}

window.MRB = {
  static: true,

  async api(path, body) {
    const py = await boot(setMsg);
    const bridge = py.pyimport('bridge');
    let out;
    try {
      out = bridge.call(path, JSON.stringify(body || {}));
    } catch (e) {
      return { ok: false, error: 'Lỗi chạy Python: ' + (e.message || e) };
    }
    try {
      return JSON.parse(out);
    } catch (e) {
      return { ok: false, error: 'Kết quả không phải JSON hợp lệ' };
    }
  },

  async browse() {
    return { ok: false, error: 'Bản chạy trong trình duyệt không duyệt ổ đĩa.' };
  },

  /* Ghi file người dùng chọn vào bộ nhớ ảo của Pyodide, rồi nhờ Python
     nhận diện loại order giống hệt bản chạy trên máy. */
  async upload(files) {
    const py = await boot(setMsg);
    setMsg(`Đang nạp ${files.length} file…`);
    for (const f of files) {
      const buf = new Uint8Array(await f.arrayBuffer());
      const name = f.name.replace(/[\\/:*?"<>|]/g, '_');
      py.FS.writeFile(WORK + '/' + name, buf, { encoding: 'binary' });
    }
    const r = await this.api('/api/files', {});
    if (!r.ok) return r;
    const known = new Set(r.exports.map(e => e.file));
    r.saved = files.map(f => ({
      file: f.name, size: f.size,
      tag: known.has(f.name) ? 'ok' : null,
    }));
    return r;
  },

  /* File xuất ra nằm trong bộ nhớ ảo -> đổi thành Blob để tải về. */
  downloadUrl(result) {
    const py = _py;
    if (!py || !result.file_path) return null;
    const bytes = py.FS.readFile(result.file_path);
    const blob = new Blob([bytes], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    return URL.createObjectURL(blob);
  },
};
