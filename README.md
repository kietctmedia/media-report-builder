# Media Report Builder

Dựng file báo cáo hiệu suất team Media (CMA) từ export Base.vn — tự động hoá đúng
quy trình đã làm tay từ T5 đến T9/2026.

**▶ Dùng ngay, không cài gì: https://kietctmedia.github.io/media-report-builder/**

Bản online chạy **hoàn toàn trong trình duyệt** (Python qua Pyodide/WebAssembly) —
không có máy chủ, **file của bạn không rời khỏi máy**. Bản chạy trên máy
(`start.bat`) mạnh hơn: quét thẳng thư mục và nhờ Excel tính sẵn công thức.

## Chạy

Nhấp đúp **`start.bat`** (hoặc `python -m app.server`). Trình duyệt tự mở
`http://127.0.0.1:8765`. Đóng cửa sổ đen là tắt app.

Không cần cài gì thêm — app chỉ dùng thư viện chuẩn của Python cộng
`openpyxl`, `xlrd`, `pywin32`, `requests` (máy đã có sẵn cả 4).

## Bốn bước

1. **Nguồn dữ liệu** — dán đường dẫn thư mục tháng (vd `E:\Báo cáo hiệu suất\2026\1026`),
   bấm *Quét thư mục*. App tự nhận loại order từ tên file và tự tìm file báo cáo
   tháng trước. Cũng có thể kéo thả file vào.
2. **Thông số** — tháng báo cáo, công chuẩn, ngày nghỉ, OT từng người.
   Công chuẩn được gợi ý = số ngày T2–T6 trừ ngày lễ, nhưng **bạn vẫn phải chốt**
   vì nó ảnh hưởng trực tiếp tới lương.
3. **Cân đối** — bảng task của từng người, sửa trực tiếp Phân loại / Standard /
   Total N, tick bỏ task. Điểm và bậc KPI cập nhật ngay. Đây là bước thay cho
   việc mở Excel sửa tay.
4. **Xuất file** — app copy file mẫu, ghi dữ liệu, nhờ Excel tính lại array formula,
   rồi chạy một loạt kiểm tra tự động. Kèm bảng phân tích theo BU.

## Logic nghiệp vụ

**Lọc:** task đã chấm điểm → người làm thuộc danh sách → ngày tạo trong tháng báo
cáo hoặc 1 tháng trước (bắt carryover) → loại mọi ID đã có trong báo cáo tháng trước.

**Carryover:** task tạo tháng trước mà hoàn thành *sau* lúc chốt báo cáo cũ thì
tính vào kỳ này; hoàn thành *trước* đó thì nhiều khả năng đã bị loại tay ở kỳ
trước nên mặc định bỏ. Mốc chốt lấy từ thời điểm hoàn thành muộn nhất trong báo
cáo cũ. Bước 3 luôn liệt kê riêng nhóm này để bạn quyết lại.

**Tính N (Total giờ):** `N = min(T, định mức)` — giờ leader phân bổ bị chặn trên
bởi rate card. Nhãn nào không có định mức (`M = Thời gian`) thì `N = T`. Riêng
task không ghi số lượng trong tên (moodboard, ebook...) thì lấy thẳng `N = T` vì
không suy ra được định mức.

**Điểm:** `Điểm thời gian = ΣN / ((Công chuẩn − Ngày nghỉ) × 8) × 10` (có thể > 10,
đây là thưởng throughput) · `Điểm chất lượng = trung bình tổng 5 tiêu chí` ·
`Điểm cuối = trung bình hai điểm trên` → tra bậc KPI.

## File cấu hình (sửa được, không cần đụng code)

| File | Dùng để |
|---|---|
| `config/rate_card.json` | Định mức giờ. Có cờ `_ap_dung_dinh_muc.bat = false` để quay về `N = T` thuần. Có `_nguong_anh_post` quyết định ranh giới *Ảnh post* (2h/ảnh) vs *Bộ ảnh social* (30p/ảnh). |
| `config/kpi_tiers.json` | Bậc KPI và số tiền. |
| `config/holidays.json` | Ngày lễ, chỉ dùng để **gợi ý** công chuẩn. |

Quy tắc phân loại theo tên task nằm ở `app/core/rules.py`, viết thành từng hàm
ngắn theo loại order — thêm luật mới chỉ là thêm một dòng `if`.

## Kiểm chứng

```
python selftest.py --build
```

Chạy lại nguyên pipeline trên dữ liệu thật tháng 9/2026 và so với báo cáo đã làm
tay. Hiện **toàn bộ kiểm tra đạt**: 132 task, Kiệt 54 task/252h, Giang 78 task/207,1h,
điểm thời gian 15,44 và 12,94 — khớp tuyệt đối, kể cả sau khi Excel tính lại.

## Base.vn API — beta, chưa kiểm chứng

Tab cuối ở bước 4. Tài liệu Workflow API của Base.vn nằm trên Postman (render bằng
JS) nên chưa dò được shape tự động, và chưa có token để thử. Nút **Thử kết nối**
gọi vài endpoint ứng viên rồi in nguyên response ra để map field.

Lấy token: `account.base.vn` → *Tích hợp với bên thứ 3 ver 2* → *Thêm application
token* (chỉ tài khoản Owner tạo được), chọn app Workflow, scope đọc job.

Khi biết shape thật, chỉ sửa `ENDPOINTS` + `FIELD_MAP` + `CUSTOM_FIELD_MAP` trong
`app/sources/base_api.py`; phần còn lại của app dùng chung một shape dữ liệu với
đường đọc file nên không phải sửa gì thêm.

## Bản online tĩnh (GitHub Pages) — đang chạy

https://kietctmedia.github.io/media-report-builder/

Không có máy chủ nào cả. Trình duyệt tải Pyodide (Python biên dịch sang
WebAssembly), cài `openpyxl` + `xlrd`, rồi chạy **đúng những module trong
`app/core` và `app/sources`** mà bản trên máy dùng. File bạn chọn chỉ nằm trong
bộ nhớ của tab, không gửi đi đâu; đóng tab là mất.

Đo trên máy thật, dữ liệu T9/2026 (5 file export, 4.655 dòng):

| | |
|---|---|
| Tải Pyodide + cài thư viện | ~2 giây (lần đầu, sau đó trình duyệt cache) |
| Đọc 5 file `.xls` + lọc + chấm điểm | ~1 giây |
| Dựng file `.xlsx` | ~1 giây |

Kết quả **trùng khít** bản chạy trên máy: 132 task, Kiệt 15,44 / Giang 12,94.

Sinh lại bản tĩnh sau khi sửa code:

```
python build_pages.py
```

Script chỉ **copy** từ `app/`, không có bản sao logic nào — sửa `rules.py` một chỗ
là cả hai bản cùng đổi.

## Bản online có máy chủ (chế độ cloud)

Dùng khi muốn host riêng thay vì GitHub Pages. Cùng một bộ code, bật bằng biến
môi trường `MRB_MODE=cloud`. Khác bản máy:

| | Bản máy (`start.bat`) | Bản online |
|---|---|---|
| Lấy file | Quét thẳng thư mục trên ổ E: | Tải file lên |
| Tính lại công thức | Có Excel → tính sẵn, tự kiểm tra | Không có Excel → số hiện khi mở file |
| Nhận file | Mở thẳng tại chỗ | Bấm nút tải về |
| Phạm vi đọc/ghi | Toàn máy | Chỉ trong `workspace/<phiên>`, tự xoá sau 6 giờ |

File xuất ra của hai bản **giống hệt nhau** — chỉ khác là bản online không nhúng
sẵn giá trị đã tính, nhưng Excel tự tính ngay lúc mở nên kết quả không đổi.

### Deploy lên Render (free)

Repo đã có sẵn `render.yaml`. Vào [dashboard.render.com](https://dashboard.render.com)
→ **New** → **Blueprint** → chọn repo này → Apply. Render đọc file đó và tự đặt
`MRB_MODE=cloud`. Khoảng 2 phút là có link `https://<tên>.onrender.com`.

Gói free ngủ sau 15 phút không ai dùng, lần vào đầu tiên chờ ~30 giây. Nếu nó ngủ
giữa chừng thì file đã tải lên có thể mất, app sẽ báo tải lại.

Host khác (Railway, Fly.io, Heroku) dùng `Procfile` có sẵn, nhớ đặt biến
`MRB_MODE=cloud`.

### Lưu ý về dữ liệu

Bản online nhận file chứa tên nhân viên, giờ làm và điểm KPI. File nằm trên máy
chủ của bên thứ ba vài giờ rồi tự xoá, và **app không có đăng nhập** — ai có link
cũng vào được. Nếu cần kín, hãy đặt mật khẩu ở tầng host (Render có Basic Auth)
hoặc chỉ dùng bản chạy trên máy.

## Vì sao bản đầy đủ phải chạy local

Hai ràng buộc bắt buộc: app phải đọc file trên ổ `E:`, và phải nhờ **Excel** tính
lại array formula (`INDEX/SMALL/IF`) trong 2 sheet Rút gọn — openpyxl chỉ ghi được
công thức chứ không tính, máy lại không có LibreOffice. Không có Excel thì file
vẫn ghi ra đúng, chỉ là số hiện lên khi bạn mở file.

## Cấu trúc

```
app/core/      rules.py (phân loại) · pipeline.py (lọc, chấm điểm, BU)
               builder.py (ghi xlsx + Excel COM) · template.py · calendar_vn.py
app/sources/   excel_export.py (.xls Base.vn) · base_api.py (beta)
app/web/       index.html · app.js · style.css
config/        rate_card.json · kpi_tiers.json · holidays.json
workspace/     nơi lưu file kéo thả và file test
selftest.py    kiểm chứng pipeline trên dữ liệu thật
```

## Bẫy đã xử lý sẵn

- File `.xls` của Base.vn lỗi OLE → đọc bằng `xlrd(..., ignore_workbook_corruption=True)`.
- Số cột khác nhau giữa các loại order → map cột **theo tên header**, không theo index.
- Console Windows cp1252 làm vỡ tiếng Việt → đã `reconfigure(encoding='utf-8')`.
- PivotTable còn sót trong template trỏ data cũ → app tự xoá trước khi ghi SUMIF.
- Nếu đang **mở file trong Excel** lúc app ghi, bản trong Excel sẽ ghi đè kết quả —
  đóng file trước khi dựng.
