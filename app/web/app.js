/* Media Report Builder — client */
'use strict';

const $  = (s, r = document) => r.querySelector(s);
const $$ = (s, r = document) => [...r.querySelectorAll(s)];
const esc = s => String(s ?? '').replace(/[&<>"]/g, c => (
  { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const h = n => (Math.round(Number(n || 0) * 100) / 100).toLocaleString('vi-VN');
const money = n => Number(n || 0).toLocaleString('vi-VN');

const S = {
  folder: '', exports: [], reports: [], template: '', templateInfo: null,
  persons: [], cfg: {}, tasks: [], edits: {}, active: '', sel: new Set(),
  stats: null, lastBuild: null, bu: null, buN: null,
};

/* ------------------------------------------------------------------ utils */
async function api(path, body) {
  const opt = body === undefined ? {} : {
    method: 'POST', headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  };
  const r = await fetch(path, opt);
  const j = await r.json().catch(() => ({ ok: false, error: 'Phản hồi không hợp lệ' }));
  return j;
}
function toast(msg, bad) {
  const t = $('#toast');
  t.textContent = msg;
  t.className = 'toast show' + (bad ? ' bad' : '');
  clearTimeout(toast._t);
  toast._t = setTimeout(() => (t.className = 'toast'), 4200);
}
function busy(on, msg) {
  $('#overlay').classList.toggle('hidden', !on);
  $('#overlayMsg').textContent = msg || '';
}
function goStep(n) {
  $$('.panel').forEach(p => p.classList.toggle('active', p.dataset.panel === String(n)));
  $$('.step').forEach(b => {
    b.classList.toggle('active', b.dataset.step === String(n));
    b.classList.toggle('done', Number(b.dataset.step) < n);
  });
  window.scrollTo({ top: 0, behavior: 'smooth' });
}
$$('.step').forEach(b => b.onclick = () => goStep(Number(b.dataset.step)));

/* =========================== BƯỚC 1: nguồn dữ liệu =========================== */
async function scan(folder) {
  busy(true, 'Đang quét thư mục…');
  const r = await api('/api/scan', { folder });
  busy(false);
  if (!r.ok) return toast(r.error, true);
  S.folder = r.folder;
  $('#folder').value = r.folder;
  setFiles(r.exports, r.reports);
}
function setFiles(exports_, reports) {
  S.exports = (exports_ || []).map(e => ({ ...e, use: true }));
  S.reports = reports || [];
  S.template = S.reports.length ? S.reports[0].path : '';
  renderFiles();
  $('#cardFiles').style.display = S.exports.length || S.reports.length ? '' : 'none';
  if (S.template) loadTemplate(S.template);
  if (!S.exports.length) toast('Không nhận ra file export nào trong thư mục này.', true);
}
const TAGS = ['QUAYCHUPMKT', 'DESIGN MKT', 'VIDEO MKT', 'RETOUCH', 'DESIGN LX', 'VIDEO LX'];

function renderFiles() {
  $('#tblExports').innerHTML = `
    <thead><tr><th class="w-s"></th><th>File export</th><th>Loại order</th><th class="num">Dung lượng</th></tr></thead>
    <tbody>${S.exports.map((e, i) => `
      <tr><td><input type="checkbox" data-ex="${i}" ${e.use ? 'checked' : ''}></td>
        <td>${esc(e.file)}</td>
        <td><select data-tag="${i}">${TAGS.map(t =>
          `<option ${t === e.tag ? 'selected' : ''}>${t}</option>`).join('')}</select></td>
        <td class="num">${(e.size / 1024 / 1024).toFixed(1)} MB</td></tr>`).join('')
      || '<tr><td colspan="4" class="hint">Chưa có file nào.</td></tr>'}</tbody>`;
  $$('[data-ex]').forEach(cb => cb.onchange = () => { S.exports[cb.dataset.ex].use = cb.checked; });
  $$('[data-tag]').forEach(sel => sel.onchange = () => { S.exports[sel.dataset.tag].tag = sel.value; });

  $('#tblReports').innerHTML = `
    <thead><tr><th class="w-s"></th><th>File báo cáo</th><th>Sửa lần cuối</th></tr></thead>
    <tbody>${S.reports.map(r => `
      <tr><td><input type="radio" name="tpl" value="${esc(r.path)}" ${r.path === S.template ? 'checked' : ''}></td>
        <td>${esc(r.file)}</td>
        <td>${new Date(r.mtime * 1000).toLocaleString('vi-VN')}</td></tr>`).join('')
      || '<tr><td colspan="3" class="hint">Không thấy file .xlsx nào có sheet Data_Tổng.</td></tr>'}</tbody>`;
  $$('input[name=tpl]').forEach(r => r.onchange = () => loadTemplate(r.value));
}

async function loadTemplate(path) {
  S.template = path;
  const r = await api('/api/template', { path });
  const box = $('#tplInfo');
  if (!r.ok) { box.className = 'callout bad'; box.textContent = r.error; return; }
  S.templateInfo = r.info;
  box.className = 'callout info';
  box.innerHTML = `<b>${esc(r.info.file)}</b> — ${r.info.task_count} task, các sheet đánh giá:
    <b>${r.info.persons.map(esc).join(', ') || '(không thấy)'}</b>.
    ${r.info.cutoff_label ? `Mốc chốt kỳ trước: <b>${r.info.cutoff_label}</b>.` : ''}
    Toàn bộ ${r.info.task_count} ID sẽ bị loại trùng.`;
  if (!S.persons.length) { S.persons = r.info.persons.slice(); renderPersons(); }
}

$('#btnScan').onclick = () => scan($('#folder').value);
$('#folder').onkeydown = e => { if (e.key === 'Enter') scan($('#folder').value); };

/* duyệt thư mục */
$('#btnBrowse').onclick = () => {
  const b = $('#browser');
  b.classList.toggle('hidden');
  if (!b.classList.contains('hidden')) browse($('#folder').value || '');
};
async function browse(path) {
  const r = await (await fetch('/api/browse?path=' + encodeURIComponent(path))).json();
  const b = $('#browser');
  if (!r.ok) { b.innerHTML = `<div class="dir">${esc(r.error)}</div>`; return; }
  const found = (r.exports || []).length + (r.reports || []).length;
  b.innerHTML =
    (r.parent !== null && r.path ? `<div class="dir up" data-p="${esc(r.parent)}">⬑ lên một cấp</div>` : '') +
    (r.path ? `<div class="dir up" data-use="${esc(r.path)}">✔ Dùng thư mục này${
      found ? ` (${found} file phù hợp)` : ''}</div>` : '') +
    r.dirs.map(d => `<div class="dir" data-p="${esc(d.path)}">📁 ${esc(d.name)}</div>`).join('');
  $$('[data-p]', b).forEach(el => el.onclick = () => browse(el.dataset.p));
  $$('[data-use]', b).forEach(el => el.onclick = () => {
    b.classList.add('hidden'); scan(el.dataset.use);
  });
}

/* upload */
const drop = $('#drop');
$('#btnPick').onclick = () => $('#fileInput').click();
$('#fileInput').onchange = e => upload(e.target.files);
['dragenter', 'dragover'].forEach(ev => drop.addEventListener(ev, e => {
  e.preventDefault(); drop.classList.add('over');
}));
['dragleave', 'drop'].forEach(ev => drop.addEventListener(ev, e => {
  e.preventDefault(); drop.classList.remove('over');
}));
drop.addEventListener('drop', e => upload(e.dataTransfer.files));

async function upload(files) {
  if (!files || !files.length) return;
  const fd = new FormData();
  [...files].forEach(f => fd.append('files', f, f.name));
  busy(true, `Đang tải ${files.length} file…`);
  const r = await (await fetch('/api/upload', { method: 'POST', body: fd })).json();
  busy(false);
  if (!r.ok) return toast(r.error, true);
  S.folder = r.folder;
  $('#folder').value = r.folder;
  const unknown = r.saved.filter(s => !s.tag).map(s => s.file);
  setFiles(r.exports, r.reports);
  toast(`Đã nhận ${r.saved.length} file${unknown.length ? ` (${unknown.length} file không nhận ra loại order)` : ''}`);
}

/* =========================== BƯỚC 2: thông số =========================== */
const now = new Date();
$('#month').innerHTML = Array.from({ length: 12 }, (_, i) =>
  `<option value="${i + 1}">Tháng ${i + 1}</option>`).join('');
const dm = new Date(now.getFullYear(), now.getMonth() - 1, 1);
$('#month').value = dm.getMonth() + 1;
$('#year').value = dm.getFullYear();

async function refreshCalendar() {
  const r = await api('/api/calendar', { year: +$('#year').value, month: +$('#month').value });
  if (!r.ok) return;
  const hol = r.holidays.map(x => `${x.date.slice(8)}/${x.date.slice(5, 7)} ${x.name}`).join(' · ');
  $('#calNote').innerHTML = `Tháng ${$('#month').value}/${$('#year').value} có
    <b>${r.weekdays}</b> ngày làm việc T2–T6${hol ? `, trùng <b>${r.holidays.length}</b> ngày lễ (${esc(hol)})` : ''}
    → gợi ý công chuẩn <b>${r.suggested}</b>.
    <span class="hint-inline">Sửa bảng ngày lễ ở <code>config/holidays.json</code>.</span>`;
  S.persons.forEach(p => {
    const el = $(`[data-cc="${CSS.escape(p)}"]`);
    if (el && !el.dataset.touched) el.value = r.suggested;
  });
  syncCfg();
}
$('#month').onchange = $('#year').onchange = refreshCalendar;

function renderPersons() {
  $('#tblPersons').innerHTML = `
    <thead><tr><th class="w-s"></th><th>Nhân sự</th><th class="w-n">Công chuẩn</th>
      <th class="w-n">Ngày nghỉ</th><th class="w-n">OT (giờ)</th><th>Mẫu số tính điểm</th></tr></thead>
    <tbody>${S.persons.map(p => `
      <tr><td><input type="checkbox" data-use-p="${esc(p)}" checked></td>
        <td><b>${esc(p)}</b></td>
        <td><input type="number" step="0.5" min="0" data-cc="${esc(p)}" value="22"></td>
        <td><input type="number" step="0.5" min="0" data-to="${esc(p)}" value="0"></td>
        <td><input type="number" step="0.5" min="0" data-ot="${esc(p)}" value="0"></td>
        <td class="hint-inline" data-denom="${esc(p)}"></td></tr>`).join('')}</tbody>`;
  $$('[data-cc],[data-to],[data-ot]').forEach(el => {
    el.oninput = () => { el.dataset.touched = '1'; syncCfg(); };
  });
  $$('[data-use-p]').forEach(cb => cb.onchange = syncCfg);
  refreshCalendar();
}
function syncCfg() {
  S.cfg = {};
  S.persons.forEach(p => {
    const on = $(`[data-use-p="${CSS.escape(p)}"]`);
    if (on && !on.checked) return;
    const cc = +($(`[data-cc="${CSS.escape(p)}"]`)?.value || 22);
    const to = +($(`[data-to="${CSS.escape(p)}"]`)?.value || 0);
    const ot = +($(`[data-ot="${CSS.escape(p)}"]`)?.value || 0);
    S.cfg[p] = { cong_chuan: cc, timeoff: to, ot };
    const d = $(`[data-denom="${CSS.escape(p)}"]`);
    if (d) d.textContent = `(${cc} − ${to}) × 8 = ${h((cc - to) * 8)} giờ`;
  });
  scheduleScore();
}

/* =========================== BƯỚC 3: phân tích & cân đối =========================== */
$('#btnAnalyse').onclick = async () => {
  syncCfg();
  const picked = S.exports.filter(e => e.use).map(e => ({ path: e.path, tag: e.tag }));
  if (!picked.length) return toast('Chưa chọn file export nào.', true);
  if (!Object.keys(S.cfg).length) return toast('Chưa chọn nhân sự nào.', true);
  busy(true, 'Đang đọc và lọc dữ liệu…');
  const r = await api('/api/analyse', {
    exports: picked, template: S.template,
    year: +$('#year').value, month: +$('#month').value,
    persons: Object.keys(S.cfg), lookback: +$('#lookback').value,
    use_cutoff: $('#useCutoff').checked,
  });
  busy(false);
  if (!r.ok) return toast(r.error, true);
  const bad = (r.read_report || []).filter(x => !x.ok);
  if (bad.length) toast(`${bad.length} file lỗi: ${bad.map(b => b.file + ' — ' + b.error).join(' | ')}`, true);
  S.tasks = r.tasks; S.stats = r.stats; S.edits = {}; S.sel.clear();
  S.active = S.active && S.tasks.some(t => t.person === S.active) ? S.active : (Object.keys(S.cfg)[0] || '');
  renderFunnel(r); renderCarry(); renderTabs(); renderTasks(); scoreNow();
  $('#outPath').value = suggestOut();
  goStep(3);
};

function suggestOut() {
  const m = +$('#month').value, y = +$('#year').value;
  const dir = S.template ? S.template.replace(/[^\\/]+$/, '') : (S.folder + '\\');
  return `${dir}MEDIA_CMA_Báo cáo hiệu suất team T${m}_${y}.xlsx`;
}

function renderFunnel(r) {
  const s = r.stats;
  const rr = (r.read_report || []).filter(x => x.ok);
  $('#funnel').innerHTML = [
    `<span>Đọc <b>${h(s.total_rows)}</b> dòng từ ${rr.length} file</span>`,
    `<span class="drop">Khác người: <b>${h(s.wrong_person)}</b></span>`,
    `<span class="drop">Ngoài kỳ ${esc(s.month_label)}${s.lookback_label ? ' + ' + esc(s.lookback_label) : ''}: <b>${h(s.out_of_month)}</b></span>`,
    `<span class="drop">Chưa chấm điểm: <b>${h(s.not_rated)}</b></span>`,
    `<span class="drop">Đã có ở báo cáo trước: <b>${h(s.duplicated)}</b></span>`,
    `<span>Còn lại: <b>${h(s.picked)}</b> task</span>`,
    s.carryover_kept ? `<span>Carryover giữ: <b>${h(s.carryover_kept)}</b></span>` : '',
    s.carryover_dropped ? `<span class="drop">Carryover tự bỏ: <b>${h(s.carryover_dropped)}</b></span>` : '',
    s.text_rating ? `<span class="drop">⚠ Chấm điểm dạng chữ: <b>${h(s.text_rating)}</b></span>` : '',
    s.no_time ? `<span class="drop">⚠ Không có giờ phân bổ: <b>${h(s.no_time)}</b></span>` : '',
  ].filter(Boolean).join('');
}

function renderCarry() {
  const rows = S.tasks.filter(t => t.carryover);
  $('#cardCarry').style.display = rows.length ? '' : 'none';
  if (!rows.length) return;
  $('#tblCarry').innerHTML = `
    <thead><tr><th class="w-s">Lấy</th><th>Task</th><th>Người</th><th class="num">Giờ</th><th>Vì sao</th></tr></thead>
    <tbody>${rows.map(t => `
      <tr><td><input type="checkbox" data-carry="${esc(t.id)}" ${cur(t).include ? 'checked' : ''}></td>
        <td>${esc(t.name)}</td><td>${esc(t.person)}</td><td class="num">${h(t.T)}</td>
        <td class="hint-inline">${esc(t.note)}</td></tr>`).join('')}</tbody>`;
  $$('[data-carry]').forEach(cb => cb.onchange = () => {
    edit(cb.dataset.carry, { include: cb.checked });
    renderTasks(); scheduleScore();
  });
}

function cur(t) { return { ...t, ...(S.edits[t.id] || {}) }; }
function edit(id, patch) { S.edits[id] = { ...(S.edits[id] || {}), ...patch }; }

function renderTabs() {
  const people = Object.keys(S.cfg);
  $('#personTabs').innerHTML = people.map(p => {
    const n = S.tasks.filter(t => t.person === p && cur(t).include).length;
    return `<button data-p="${esc(p)}" class="${p === S.active ? 'active' : ''}">${esc(p)} <span class="badge">${n}</span></button>`;
  }).join('');
  $$('#personTabs button').forEach(b => b.onclick = () => {
    S.active = b.dataset.p; S.sel.clear(); renderTabs(); renderTasks();
  });
}

function visibleTasks() {
  const q = ($('#search').value || '').toLowerCase().trim();
  return S.tasks.filter(t => t.person === S.active &&
    (!q || t.name.toLowerCase().includes(q) || t.L.toLowerCase().includes(q)));
}

function renderTasks() {
  const rows = visibleTasks();
  $('#tblTasks').innerHTML = `
    <thead><tr>
      <th class="w-s"><input type="checkbox" id="selAll"></th>
      <th class="w-s">STT</th><th>Tên task</th><th>BU</th>
      <th class="num w-s">Giờ</th><th class="num w-s">CL</th>
      <th class="w-l">Phân loại (L)</th><th class="w-m">Standard (M)</th><th class="w-n">Total N</th>
    </tr></thead>
    <tbody>${rows.map((t0, i) => {
      const t = cur(t0);
      const changed = t.N !== t0.sN || t.L !== t0.sL || t.M !== t0.sM;
      return `<tr class="${t.include ? '' : 'excluded'} ${changed ? 'edited' : ''}" data-row="${esc(t.id)}">
        <td class="keep"><input type="checkbox" data-sel="${esc(t.id)}" ${S.sel.has(t.id) ? 'checked' : ''}></td>
        <td class="keep num">${i + 1}</td>
        <td title="${esc(t.name)}">${esc(t.name)}${t.carryover ? ' <span class="badge carry">carryover</span>' : ''}</td>
        <td>${esc(t.bu)}</td>
        <td class="num">${h(t.T)}</td>
        <td class="num">${h(t.quality)}</td>
        <td class="keep"><input type="text" data-l="${esc(t.id)}" value="${esc(t.L)}"></td>
        <td class="keep"><input type="text" data-m="${esc(t.id)}" value="${esc(t.M)}"></td>
        <td class="keep ncell"><input type="number" step="0.1" min="0" data-n="${esc(t.id)}" value="${t.N}"></td>
      </tr>`;
    }).join('') || '<tr><td colspan="9" class="hint">Không có task nào khớp.</td></tr>'}</tbody>`;

  $$('[data-sel]').forEach(cb => cb.onchange = () => {
    cb.checked ? S.sel.add(cb.dataset.sel) : S.sel.delete(cb.dataset.sel);
    updateSelCount();
  });
  $('#selAll').onchange = e => {
    rows.forEach(t => e.target.checked ? S.sel.add(t.id) : S.sel.delete(t.id));
    renderTasks();
  };
  $$('[data-n]').forEach(el => el.onchange = () => {
    edit(el.dataset.n, { N: +el.value }); renderTasks(); scheduleScore();
  });
  $$('[data-l]').forEach(el => el.onchange = () => { edit(el.dataset.l, { L: el.value }); scheduleScore(); });
  $$('[data-m]').forEach(el => el.onchange = () => { edit(el.dataset.m, { M: el.value }); scheduleScore(); });
  updateSelCount();
}
function updateSelCount() {
  $('#selCount').textContent = S.sel.size ? `đã chọn ${S.sel.size} dòng` : 'chưa chọn dòng nào';
}
$('#search').oninput = renderTasks;

$$('[data-bulk]').forEach(b => b.onclick = () => {
  const ids = S.sel.size ? [...S.sel] : visibleTasks().map(t => t.id);
  if (!ids.length) return;
  const byId = Object.fromEntries(S.tasks.map(t => [t.id, t]));
  ids.forEach(id => {
    const t = byId[id]; if (!t) return;
    switch (b.dataset.bulk) {
      case 'dinhmuc': edit(id, { N: t.sN }); break;
      case 'T':       edit(id, { N: t.T }); break;
      case 'reset':   edit(id, { N: t.sN, L: t.sL, M: t.sM, include: t.include }); break;
      case 'exclude': edit(id, { include: false }); break;
      case 'include': edit(id, { include: true }); break;
    }
  });
  toast(`Đã áp dụng cho ${ids.length} dòng${S.sel.size ? '' : ' (toàn bộ dòng đang hiện)'}`);
  renderTabs(); renderTasks(); renderCarry(); scheduleScore();
});

let scoreTimer;
function scheduleScore() { clearTimeout(scoreTimer); scoreTimer = setTimeout(scoreNow, 220); }

async function scoreNow() {
  if (!S.tasks.length || !Object.keys(S.cfg).length) return;
  const r = await api('/api/score', { edits: S.edits, persons_cfg: S.cfg });
  if (!r.ok) return;
  S.bu = r.bu; S.buN = r.bu_n;
  $('#scorebar').innerHTML = Object.entries(r.scores).map(([p, s]) => `
    <div class="scorecard">
      <h3>${esc(p)}</h3>
      <div><span class="big">${s.final_score}</span>
        <span class="tier ${s.money ? '' : 'low'}">${esc(s.tier)} · ${money(s.money)}đ</span></div>
      <dl>
        <dt>Điểm thời gian</dt><dd>${s.time_score}</dd>
        <dt>Điểm chất lượng</dt><dd>${s.quality_score}</dd>
        <dt>Số task</dt><dd>${s.tasks}</dd>
        <dt>Tổng giờ tính điểm</dt><dd>${h(s.total_n)} h</dd>
        <dt>Mẫu số</dt><dd>${h(s.denominator)} h</dd>
      </dl>
    </div>`).join('');
  renderBU();
}

/* =========================== BƯỚC 4: xuất file =========================== */
$('#btnBackTo2').onclick = () => goStep(2);
$('#btnToExport').onclick = () => { if (!$('#outPath').value) $('#outPath').value = suggestOut(); goStep(4); };

$('#btnBuild').onclick = () => doBuild(false);

async function doBuild(overwrite) {
  if (!S.tasks.length) return toast('Chưa có dữ liệu — chạy bước phân tích trước.', true);
  busy(true, 'Đang dựng file và nhờ Excel tính lại…');
  const r = await api('/api/build', {
    edits: S.edits, template: S.template, out_path: $('#outPath').value,
    persons_cfg: S.cfg, overwrite,
  });
  busy(false);
  if (!r.ok && r.exists) {
    if (confirm('File đã tồn tại:\n' + r.path + '\n\nGhi đè?')) return doBuild(true);
    return;
  }
  if (!r.ok) {
    $('#buildLog').innerHTML = `<div class="callout bad"><b>Lỗi:</b> ${esc(r.error)}
      ${r.trace ? `<pre class="probe">${esc(r.trace)}</pre>` : ''}</div>`;
    return toast(r.error, true);
  }
  S.lastBuild = r;
  renderResult(r);
  toast('Đã xuất file xong.');
}

function renderResult(r) {
  $('#buildLog').innerHTML = `<div class="callout info">Đã ghi
    <b>${r.layout_warnings.length ? '' : ''}${esc(r.out_path)}</b></div>`;
  $('#cardChecks').style.display = '';
  const rec = r.recalc || {};
  const master = rec.master && Object.keys(rec.master).length ? rec.master : null;

  let html = '';
  if (!rec.available) {
    html += `<div class="callout warn"><b>Chưa tính lại được bằng Excel</b> — ${esc(rec.reason || '')}.
      File vẫn đúng công thức; mở bằng Excel là số sẽ hiện ra.</div>`;
  }
  html += '<div class="scorebar">' + Object.entries(r.preview).map(([p, s]) => {
    const m = master && master[p];
    return `<div class="scorecard">
      <h3>${esc(p)} ${m ? '<span class="badge ok">Excel xác nhận</span>' : ''}</h3>
      <div><span class="big">${m ? m.final_score : s.final_score}</span>
        <span class="tier ${(m ? m.money : s.money) ? '' : 'low'}">${esc(m ? m.tier : s.tier)} · ${money(m ? m.money : s.money)}đ</span></div>
      <dl><dt>Điểm thời gian</dt><dd>${m ? m.time_score : s.time_score}</dd>
          <dt>Điểm chất lượng</dt><dd>${m ? m.quality_score : s.quality_score}</dd>
          <dt>Số task</dt><dd>${s.tasks}</dd>
          <dt>Tổng giờ tính điểm</dt><dd>${h(s.total_n)} h</dd></dl></div>`;
  }).join('') + '</div>';

  const checks = r.checks || [];
  const nbad = checks.filter(c => !c.ok).length;
  html += `<h3 style="margin:16px 0 4px">Kiểm tra tự động
    <span class="badge ${nbad ? 'bad' : 'ok'}">${checks.length - nbad}/${checks.length} đạt</span></h3>`;
  html += checks.map(c => `<div class="checkrow">
      <span class="mark ${c.ok ? 'ok' : 'bad'}">${c.ok ? '✓' : '!'}</span>
      <span>${esc(c.name)}</span><span class="det">${esc(c.detail || '')}</span></div>`).join('');

  (r.layout_warnings || []).forEach(w => {
    html += `<div class="callout warn">${esc(w)}</div>`;
  });
  html += `<div class="row" style="margin-top:16px">
      <button class="btn" id="btnOpenFile">Mở file</button>
      <button class="btn ghost" id="btnOpenDir">Mở thư mục chứa file</button></div>`;
  $('#resultBox').innerHTML = html;
  $('#btnOpenFile').onclick = () => api('/api/open', { path: r.out_path });
  $('#btnOpenDir').onclick = () => api('/api/open', { path: r.out_path, reveal: true });
}

/* BU */
$$('input[name=bumetric]').forEach(el => el.onchange = renderBU);
function renderBU() {
  const metric = $$('input[name=bumetric]').find(e => e.checked)?.value || 'T';
  const data = metric === 'T' ? S.bu : S.buN;
  if (!data || !data.rows.length) { $('#buBox').innerHTML = '<p class="hint">Chưa có dữ liệu.</p>'; return; }
  const people = Object.keys(S.cfg);
  $('#buBox').innerHTML = `<table class="tbl">
    <thead><tr><th>BU</th><th class="num">Task</th><th class="num">Giờ</th><th class="num">%</th>
      ${people.map(p => `<th>${esc(p)}</th>`).join('')}</tr></thead>
    <tbody>${data.rows.map(row => `<tr>
      <td><b>${esc(row.bu)}</b></td><td class="num">${row.tasks}</td>
      <td class="num">${h(row.hours)}</td><td class="num">${row.pct}%</td>
      ${people.map(p => {
        const v = row.per_person[p];
        return `<td class="num hint-inline">${v ? `${v.tasks} task · ${h(v.hours)}h` : '—'}</td>`;
      }).join('')}</tr>`).join('')}
    <tr><td><b>TỔNG</b></td><td class="num"><b>${data.total_tasks}</b></td>
      <td class="num"><b>${h(data.total_hours)}</b></td><td class="num">100%</td>
      ${people.map(() => '<td></td>').join('')}</tr></tbody></table>`;
}

/* Base API probe */
$('#btnProbe').onclick = async () => {
  busy(true, 'Đang gọi thử API Base.vn…');
  const r = await api('/api/base/probe', { token: $('#baseToken').value, extra: $('#baseExtra').value });
  busy(false);
  const out = $('#probeOut');
  out.classList.remove('hidden');
  if (!r.ok) { out.textContent = r.error; return; }
  out.textContent = r.results.map(x => {
    const lines = [`${x.method} ${x.url}`, `  HTTP ${x.status ?? '—'}${x.error ? '  ' + x.error : ''}`];
    if (x.json_keys) lines.push(`  keys: ${JSON.stringify(x.json_keys)}`);
    if (x.list_key) lines.push(`  danh sách ở key "${x.list_key}", ${x.count} phần tử`);
    if (x.sample_fields) lines.push(`  field của 1 job: ${JSON.stringify(x.sample_fields)}`);
    if (x.sample) lines.push(`  mẫu: ${x.sample}`);
    else if (x.raw) lines.push(`  raw: ${x.raw.slice(0, 600)}`);
    return lines.join('\n');
  }).join('\n\n');
};

/* khởi động */
(async () => {
  goStep(1);
  const st = await (await fetch('/api/state')).json();
  if (st.ok && st.folder) { $('#folder').value = st.folder; }
  refreshCalendar();
})();
