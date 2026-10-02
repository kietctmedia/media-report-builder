/* Transport cho bản chạy trên máy: gọi HTTP tới server Python. */
'use strict';

window.MRB = {
  static: false,

  async api(path, body) {
    const r = await fetch(path, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body || {}),
    });
    return r.json().catch(() => ({ ok: false, error: 'Phản hồi không hợp lệ từ máy chủ' }));
  },

  async browse(path) {
    const r = await fetch('/api/browse?path=' + encodeURIComponent(path || ''));
    return r.json().catch(() => ({ ok: false, error: 'Không duyệt được thư mục' }));
  },

  async upload(files) {
    const fd = new FormData();
    files.forEach(f => fd.append('files', f, f.name));
    const r = await fetch('/api/upload', { method: 'POST', body: fd });
    return r.json().catch(() => ({ ok: false, error: 'Tải file lên thất bại' }));
  },
};
