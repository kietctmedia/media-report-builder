# Bàn giao quy trình: Báo cáo hiệu suất team Media (CMA)

> Tài liệu này dành cho **Claude chạy trên máy khác** tiếp nhận công việc.
> Đọc hết trước khi làm. Mọi con số và quy tắc ở đây đều đã được kiểm chứng trên
> dữ liệu thật từ tháng 5 đến tháng 9/2026.

**Link gốc (fetch được):**
`https://raw.githubusercontent.com/kietctmedia/media-report-builder/main/HANDOFF.md`

---

## 1. Việc này là gì

Mỗi tháng dựng **một file Excel báo cáo hiệu suất** cho team Media của SAPP (sản
phẩm CMA), đánh giá 2 nhân sự:

| Người | Vị trí | Làm gì |
|---|---|---|
| **Nguyễn Hương Giang** | A2 | Design |
| **Cao Tuấn Kiệt** | A3 | Design + Video |

Kết quả ra **mức KPI và số tiền thưởng**, nên mọi con số đều ảnh hưởng tới lương —
tuyệt đối không đoán bừa những tham số phải hỏi (xem §7).

Thư mục làm việc: `E:\Báo cáo hiệu suất\2026\<MMDD>\`
(ví dụ tháng 10/2026 → `E:\Báo cáo hiệu suất\2026\1026\`).

---

## 2. Nguyên liệu đầu vào

User tự export từ Base.vn và bỏ vào thư mục tháng. **Số lượng file thay đổi theo
tháng** — đừng hardcode, cứ liệt kê thư mục rồi nhận dạng theo tên file:

| Tên file chứa | Loại order (tag) |
|---|---|
| `1-1-order-quay-chup-mkt` | `QUAYCHUPMKT` |
| `1-2-order-design-mkt` | `DESIGN MKT` |
| `1-3-order-edit-video-mkt` | `VIDEO MKT` |
| `1-4-order-retouch-sua-anh` | `RETOUCH` |
| `2-2-order-design-lx` | `DESIGN LX` |
| `2-3-order-edit-video-lx` | `VIDEO LX` |

Tháng 8 có 6 file, tháng 9 chỉ có 5 (không có `2-3`). Nếu thiếu file mà tháng đó
thực tế có task, phải hỏi user chứ đừng im lặng bỏ qua.

Ngoài ra cần **file báo cáo của tháng trước** (`.xlsx`) — vừa làm khuôn, vừa là
nguồn ID để loại trùng.

### Đọc file export

```python
import xlrd
book = xlrd.open_workbook(path, ignore_workbook_corruption=True)
sheet = book.sheet_by_name('Tất cả')
```

- File `.xls` của Base.vn **bị lỗi OLE** → pandas và openpyxl đều fail. Bắt buộc
  dùng `xlrd` với `ignore_workbook_corruption=True`.
- **Số cột KHÔNG cố định** giữa các loại order (bản `edit-video-lx` từng có 30 cột
  do chèn 2 cột học phần ACCA/CFA). → **Map cột theo TÊN HEADER**, không theo index.
- 5 cột chấm điểm = 5 cột **ngay sau** `Thời gian hoàn thành task` (tên tiêu chí
  khác nhau theo từng loại order). Sau đó 1 cột feedback, rồi tới cột `ID`.
  Dùng `assert` kiểm tra giả định này.
- Mọi giá trị đều là **chuỗi**: ngày `'DD/MM/YYYY HH:MM'`, điểm `'2'` / `'1.5'`.

---

## 3. Cấu trúc file báo cáo (5 sheet)

### `Data_Tổng` — dữ liệu gốc

| Cột | Nội dung |
|---|---|
| A | Tag loại order |
| B–R | 17 field thô từ Base.vn (Ngày tạo, Người tạo, Giao cho, Tên nhiệm vụ, Mô tả, Người theo dõi, Giai đoạn, Trạng thái, Nhãn, Lý do thất bại, Thời gian hoàn thành, Thời gian cập nhật, Ngày muốn nhận, Brief, Phòng ban, Kích thước, Mẫu ấn phẩm) |
| S | Người làm task |
| T | **Thời gian leader phân bổ** (giờ) — số gốc quan trọng nhất |
| U | Thời gian hoàn thành task |
| V–Z | 5 tiêu chí chấm điểm (mỗi cái 1.5 hoặc 2) |
| AA | Feedback |
| AB | ID task |
| AC | `=SUM(V{r}:Z{r})` — tổng 5 tiêu chí |
| AD | Nhóm (Quay/Chụp · Design · Video social · Retouch · Video SX) |
| AE / AF | Khối SUMIF: tổng giờ phân bổ theo người + tổng chung |

Dòng dữ liệu bắt đầu từ **dòng 2**.

Thứ tự ghi: **gom theo tag** (theo thứ tự bảng ở §2), trong mỗi tag **giữ nguyên
thứ tự gốc của file export**. Thứ tự này quan trọng — xem §9.

### `Kiệt_Rút gọn` và `Giang_Rút gọn`

Dòng dữ liệu bắt đầu từ **dòng 3**.

| Cột | Nội dung |
|---|---|
| A | STT (1..n, liền mạch) |
| B–I, K | **Array formula** kéo từ Data_Tổng (xem bên dưới) |
| J | `="https://workflow.base.vn/job/" & K{r}` |
| L | **Phân loại** (nhãn mô tả loại việc) |
| M | **Standard** (định mức tham chiếu, vd `2h/ảnh`) |
| N | **Total (Giờ)** — số dùng để tính điểm |

Array formula (phải ghi **đúng** dạng này, dùng `ArrayFormula` 1 ô của openpyxl):

```
=IFERROR(INDEX(Data_Tổng!$E:$E, SMALL(IF(Data_Tổng!$S$2:$S$997="Kiệt", ROW(Data_Tổng!$S$2:$S$997)), ROW(A1))), "")
```

Map cột nguồn: `B←E` (tên task), `C←V`, `D←W`, `E←X`, `F←Y`, `G←Z` (5 tiêu chí),
`H←AC` (tổng chất lượng), `I←T` (giờ phân bổ), `K←AB` (ID).
Phần `ROW(A1)` tăng dần `A1, A2, A3…` theo từng dòng.

**Khối cấu hình** nằm ở cột O/P/Q (15/16/17) — **O là NHÃN, P là GIÁ TRỊ**:

| Ô | Nội dung |
|---|---|
| O3 / P3 | `Công chuẩn` / số |
| O4 / P4 | `TIMEOFF` / số ngày nghỉ |
| O5 / P5 | `OT` / số giờ |
| O8 / P8 / Q8 | `Điểm thời gian` / `Điểm chất lượng` / `Điểm KPI` |
| O9 | `=((SUM(N3:N{last}))/((P3-P4)*8))*10` |
| P9 | `=AVERAGE(H3:H{last})` |
| Q9 | `=AVERAGE(O9:P9)` |

> ⚠️ Ghi nhầm nhãn sang cột P → O9 thành `text − text` = `#VALUE!`.

### `Master` — tổng hợp

Dòng 2 = Giang, dòng 3 = Kiệt. Cột E/F kéo điểm từ sheet Rút gọn
(`=ROUND('Giang_Rút gọn'!O9,2)`), G = trung bình, H = bậc (`IFS`), I = tiền.

### `Cơ chế Standard Hours` — rate card, giữ nguyên, không đụng vào.

---

## 4. Quy trình 7 bước

1. Liệt kê thư mục tháng → nhận dạng file export + file báo cáo tháng trước.
2. Đọc toàn bộ dòng từ các file export.
3. **Lọc + dedup** (§5).
4. **Phân loại L/M/N** (§6).
5. **Hỏi user** công chuẩn / ngày nghỉ / OT (§7).
6. Dựng file `.xlsx` từ template tháng trước, rồi **nhờ Excel tính lại** (§9).
7. **Chạy kiểm tra** (§10) rồi báo cáo kết quả kèm phân tích BU (§11).

---

## 5. Logic lọc và loại trùng

Giữ lại task thoả **đồng thời**:

1. **Đã được chấm điểm** — cả 5 cột tiêu chí đều có số.
2. **Người làm task** ∈ {Kiệt, Giang}.
3. **Ngày tạo** ∈ {tháng báo cáo} ∪ {tháng liền trước} — lookback 1 tháng để bắt
   task làm vắt sang tháng sau (carryover).
4. **Trừ toàn bộ ID** đã có trong báo cáo tháng trước.

### Bẫy carryover — quan trọng

Sau khi dedup, các task **tạo ở tháng lookback** chia 2 loại:

- Hoàn thành **SAU** lúc chốt báo cáo tháng trước → carryover thật, **lấy**.
- Hoàn thành **TRƯỚC** đó → task này đã từng xuất hiện ở kỳ trước và **bị user
  loại tay**, không được để nó quay lại. **Mặc định bỏ**, nhưng vẫn liệt kê ra
  cho user xem.

Mốc chốt = **thời điểm hoàn thành muộn nhất** trong Data_Tổng của báo cáo cũ.

Quy tắc này đã kiểm chứng ngược trên 2 kỳ: T7→T8 lọc đúng 7 task re-edit/retouch
trùng mà user đã bỏ tay; T8→T9 giữ đúng 11 task carryover thật.

---

## 6. Phân loại L / M / N

### Quy tắc tính N — chốt từ T8/2026

```
N = min(T, định mức)
```

trong đó `T` = giờ leader phân bổ (cột T của Data_Tổng). Nhãn nào không có định
mức (`M = 'Thời gian'`) thì `N = T`.

> Trước T8 từng dùng `N = T` thuần. Từ T8 user tự tay hạ giờ những task vượt định
> mức → quy tắc `min` tái lập đúng 146/153 dòng file T8 của user.

### Bảng quy tắc

**DESIGN MKT / DESIGN LX** (xét theo thứ tự này):

| Điều kiện trong tên task | L | M | N |
|---|---|---|---|
| có `resize` | `Ảnh social_Resize_{c} ảnh` | `10p/ảnh` | `T` |
| có `design system` | `Design System` | `Thời gian` | `T` |
| có `sửa` / `chỉnh sửa` / `tách nền` / `tách element` | `Ảnh social_Sửa` | `Thời gian` | `T` |
| có `kv` (ranh giới từ) hoặc `key visual` | `KV` | `4h` | `min(T, 4)` |
| có `{n} section` | `Landingpage_{n} section` | `40p/section` | `min(T, n×⅔)` |
| **không** ghi số ảnh | `Ảnh post` | `2h/ảnh` | `T` |
| `1 ảnh` | `Ảnh post` | `2h/ảnh` | `min(T, 2)` |
| `{c} ảnh` và `T/c ≥ 2` | `Ảnh post` | `2h/ảnh` | `min(T, c×2)` |
| `{c} ảnh` và `T/c < 2` | `Bộ ảnh social_{c} ảnh` | `30p/ảnh` | `min(T, c×0.5)` |

**QUAYCHUPMKT:**

| Điều kiện | L | M | N |
|---|---|---|---|
| có `retouch` | `Retouch` | `Thời gian` | `T` |
| có `SC`+`KG` và `T ≥ 6` | `Chụp lớp SC + quay highlight + quay phỏng vấn` | `Thời gian` | `T` |
| có `lớp`/`trial`/`khai giảng`/`masterclass` + `quay` | `Quay lớp học` | `3h` | `min(T, 3)` |
| có `lớp`/`trial`/`khai giảng`/`masterclass` | `Chụp lớp học` | `Thời gian` | `T` |
| có `chụp` | `Chụp sự kiện` | `Thời gian` | `T` |
| có `quay` | `Quay sự kiện` | `Thời gian` | `T` |

**VIDEO MKT:**

| Điều kiện | L | M | N |
|---|---|---|---|
| có `retouch` | `Retouch` | `Thời gian` | `T` |
| có `claude`/`premiere`/`setup`/`skill`/`agent` | `Video social_Setup AI` | `Thời gian` | `T` |
| có `sửa` | `Video social_Sửa` | `Thời gian` | `T` |
| có `xào lại`, hoặc là bản ver đánh số của 1 video gốc | `Video social_Ver` | `Thời gian` | `T` |
| có `animation` | `Video animation` | `6h` | `min(T, 6)` |
| còn lại | `Video social` | `6h` | `min(T, 6)` |

**RETOUCH:** luôn `Retouch` / `Thời gian` / `T`.

**VIDEO LX:** có `convert` → `Convert record`/`Thời gian`/`T`;
`T < 3` → `Video SX_nhỏ lẻ`/`Thời gian`/`T`;
còn lại → `Video SX_{round(T/3)} buổi`/`3h`/`T`.

### Hai bẫy regex

- `KV` phải dùng `(?:^|[\s_])kv(?:[\s_]|$)` — dấu `_` phá `\b` của regex.
- `Landingpage` **chỉ** trigger bởi chữ `section`, KHÔNG phải chữ `landing`.

### Chỗ không có luật — phải nói với user

Ranh giới **`Ảnh post` (2h/ảnh) vs `Bộ ảnh social` (30p/ảnh)** dùng ngưỡng
`T/c ≥ 2`, nhưng đây là **đòn bẩy cân điểm thủ công của user**, không suy ra được:

> Ở T8 có 21 task ads "3 ảnh, T=6h" gần như giống hệt nhau. User tự tay chuyển
> **7 task** sang `Bộ ảnh social` (N: 6h → 1.5h) và giữ **14 task** ở `Ảnh post`
> (N=6h). Không có dấu hiệu nào trong dữ liệu phân biệt được.

→ Cứ build theo ngưỡng `T/c ≥ 2`, rồi **chủ động nói cho user biết** đây là chỗ
để họ cân điểm nếu muốn.

---

## 7. BẮT BUỘC hỏi user mỗi tháng

Ba tham số này ảnh hưởng trực tiếp tới lương, **không được đoán**:

1. **Công chuẩn** — số ngày làm việc chuẩn. Gợi ý = số ngày T2–T6 trong tháng trừ
   ngày nghỉ lễ, nhưng user chốt.
2. **TIMEOFF** — số ngày nghỉ của **từng người**.
3. **OT** — số giờ làm thêm của **từng người**.

Lịch sử: T6 = 22 (cả hai nghỉ 2) · T7 = 23 (nghỉ 0) · T8 = 21 (Kiệt nghỉ 1) ·
T9 = 20 (22 ngày trừ 2 ngày lễ 2/9, nghỉ 0).

> Nếu user không trả lời: vẫn **làm xong toàn bộ** với giả định ghi rõ ràng, nói
> thẳng đó là giả định, và chỉ ra rằng đổi tham số chỉ cần sửa 3 ô P3/P4/P5 rồi
> tính lại. Đừng dừng lại chờ.

---

## 8. Chấm điểm

```
Điểm thời gian   = ΣN / ((Công chuẩn − TIMEOFF) × 8) × 10
Điểm chất lượng  = trung bình (tổng 5 tiêu chí)
Điểm cuối        = trung bình (Điểm thời gian, Điểm chất lượng)
```

**Điểm thời gian CÓ THỂ > 10** — đó là thưởng throughput, không phải lỗi
(T5 Kiệt = 15.9, T9 Kiệt = 15.44). Đừng "sửa" cho nó về 10.

Bậc KPI theo Điểm cuối:

| Điểm cuối | Bậc | Tiền |
|---|---|---|
| ≥ 8.9 | MỨC 3 | 3.000.000 |
| ≥ 8.4 | MỨC 2 | 2.500.000 |
| ≥ 8.0 | MỨC 1 | 2.000.000 |
| < 8.0 | Không đạt KPI | 0 |

---

## 9. Dựng file

### Cách 1 — dùng app có sẵn (khuyến nghị)

```bash
git clone https://github.com/kietctmedia/media-report-builder
cd media-report-builder
start.bat          # hoặc: python -m app.server
```

Mở `http://127.0.0.1:8765`, đi 4 bước: chọn thư mục → nhập thông số → cân đối →
xuất file. Toàn bộ logic ở §5–§8 đã nằm trong đó.

Bản chạy trong trình duyệt (không cài gì):
**https://kietctmedia.github.io/media-report-builder/**

Kiểm chứng app còn đúng: `python selftest.py --build` (chạy lại trên dữ liệu
T9/2026 thật, phải ra 132 task · Kiệt 15.44 · Giang 12.94).

### Cách 2 — viết script tay

1. `shutil.copyfile(template_thang_truoc, file_moi)`
2. openpyxl: xoá sạch Data_Tổng từ dòng 2 (cột 1–32) và 2 sheet Rút gọn từ dòng 3
   (cột 1–14), rồi ghi lại.
3. Cập nhật range: `SUMIF($S$2:$S${last})`, `SUM(N3:N{last})`, `AVERAGE(H3:H{last})`.
4. **Tính lại bằng Excel COM** (xem dưới).

### Vì sao phải nhờ Excel

2 sheet Rút gọn dùng array formula `INDEX/SMALL/IF`. openpyxl chỉ **ghi** được
công thức, không tính. Máy không có LibreOffice, và `recalc.py` của skill `xlsx`
dùng **unix socket** nên **chết trên Windows**. → Dùng Excel COM:

```python
import win32com.client as w32
app = w32.DispatchEx('Excel.Application')
app.Visible = False; app.DisplayAlerts = False
wb = app.Workbooks.Open(path)
app.CalculateFullRebuild(); app.CalculateFull()
wb.Save(); wb.Close(SaveChanges=True); app.Quit()
```

- **Chỉ mở MỘT phiên** `DispatchEx` cho cả ghi + tính + đọc kiểm tra. Mở 2 phiên
  tuần tự trong cùng script Python → **segfault**.
- Dọn tiến trình mồ côi: `taskkill /F /IM EXCEL.EXE`.
- Không có Excel cũng không sao: file vẫn đúng công thức, số hiện khi mở lên.

---

## 10. Kiểm tra bắt buộc trước khi báo cáo xong

| Kiểm tra | Đạt khi |
|---|---|
| Số dòng ghi vào file = số task đã chọn | bằng nhau |
| Không có ID trùng | 0 trùng |
| Mỗi dòng Rút gọn kéo đúng ID theo thứ tự Data_Tổng | 0 lệch |
| L/M/N canh đúng dòng (so lại theo ID) | 0 lệch |
| Không có ô trống trong vùng dữ liệu (cột B, H, I, K, N) | 0 ô |
| Không có lỗi công thức (`#N/A`, `#VALUE!`…) | 0 lỗi |
| Dòng ngay sau vùng dữ liệu phải rỗng | rỗng |
| SUMIF ở Data_Tổng khớp tổng giờ phân bổ | khớp |
| Điểm Excel tính ra khớp bản tính tay | lệch < 0.02 |

**Báo cáo trung thực**: nếu có kiểm tra nào không đạt thì nói rõ, đừng giấu.

---

## 11. Phân tích theo BU (user tháng nào cũng hỏi)

Tách theo **tiền tố tên task**, metric = **Σ giờ leader phân bổ (T)**.
Xét theo đúng thứ tự này:

1. **SC** — có `_SC_`, `SC{số}`, `ShortCourse`, `Short course`
   (3 biến thể: `0926_CMA_SC_...`, `0926_ShortCourse_...` không có `CMA` ở giữa,
   và `ShortCourse CMA`)
2. **CGMA** — có `_CGMA_`
3. **B2B** — có `_B2B_`
4. **CEO** — có `_CEO_` hoặc `TOFU_CMC {số}` (lớp CEO Master Class)
5. **FASIS** — có `fasis`
6. **HO/Nội bộ** — có `office tour` hoặc `_HO_`
7. **CMA** — có `_CMA_`
8. còn lại → `Khác`

> ⚠️ CGMA và B2B **phải xét trước CEO**, nếu không task
> `0926_B2B_Book lịch quay Event CEO Forum` bị tính nhầm vào CEO.

---

## 12. Bẫy đã gặp — đọc kỹ

### Về file

- **User đang mở file trong Excel** lúc mình ghi qua COM → bản in-memory của họ
  **ghi đè** thay đổi của mình khi họ lưu. **Luôn dặn user đóng file trước.**
  File mới (chưa ai mở) thì an toàn.
- **PivotTable sót trong template**: `Data_Tổng!AE2:AF5` ở một số bản cũ là
  PivotTable thật. Copy sang tháng mới thì pivot cache trỏ dữ liệu CŨ → hiện số
  SAI, và ghi SUMIF bằng openpyxl sẽ bị Excel ghi đè lại khi mở. → Phải **xoá
  pivot qua COM** (`ws.PivotTables(i).TableRange2.Clear()`) rồi mới ghi SUMIF.
- **LUÔN backup** file user đã sửa tay trước khi ghi đè.

### Về môi trường

- **Console Windows cp1252** làm crash khi `print` tiếng Việt →
  `sys.stdout.reconfigure(encoding='utf-8')` ở đầu mọi script.
- **Bash tool nuốt backslash trong heredoc** → nội dung có `\` phải dùng Write tool.

### Về dữ liệu

- **Widget chấm điểm của Base.vn đã đổi sang dạng CHỮ** cho một số loại order
  (`'Đạt'`, `'Xuất sắc vượt mong đợi'`, `'10 điểm (Đúng deadline)'`…). Tính tới T9,
  task của Kiệt/Giang **vẫn chấm bằng số** nên parse bình thường. **Nhưng phải
  kiểm tra lại mỗi tháng** — nếu chuyển sang chữ thì phải map chữ → số (mỗi tiêu
  chí tối đa 2 điểm) trước khi tính.
- Loại order mới xuất hiện bất ngờ (T8 có `DESIGN LX` lần đầu, T9 mất `VIDEO LX`).
  BU mới cũng vậy (T9 lần đầu có tuyến **CEO**, chiếm ngay 17% giờ).

### Về cách user làm việc

User hay **"cân đối" tay** sau khi nhận file: sửa L/M/N để cân điểm 2 người, và
**xoá trắng cột N** của những task muốn loại khỏi điểm.

Khi user bảo *"điều chỉnh file tổng + KPI theo 2 sheet"*:

1. Đọc L/M/N tay theo **ID** + tập ID giữ lại của mỗi người.
2. Rebuild Data_Tổng = bỏ ID bị loại, **GIỮ NGUYÊN THỨ TỰ** (không thì sheet của
   người kia bị xô lệch).
3. Rebuild sheet Rút gọn: STT liền mạch 1..n, array formula `ROW(A1..An)`, gán lại
   L/M/N **theo ID**.
4. Tính lại bằng COM.
5. Verify: so L/M/N từng ID với bản backup = 0 lệch.

Nếu user muốn **bản riêng không ghi đè** → đặt tên `..._Ver2.xlsx`.

---

## 13. Lịch sử số liệu (để sanity-check)

| Tháng | Công chuẩn | Task | Kiệt | Giang | Ghi chú |
|---|---|---|---|---|---|
| T5 | — | 117 | 15.9 | — | 5 task tạo T4 |
| T6 | 22 (nghỉ 2/2) | 134 | — | — | 3 task carryover |
| T7 | 23 (nghỉ 0) | 147 → **139** | 13.72 → 10.83 | 9.45 → 10.71 | user bỏ 7-8 task re-edit trùng |
| T8 | 21 (Kiệt nghỉ 1) | 153 | 14.47 → 14.28 | 16.98 → 15.10 | user hạ N 7 task ads |
| T9 | 20 (2 ngày lễ 2/9) | 132 → **117** | 15.44 → 10.63 | 12.94 → 10.69 | user bỏ 15 task video |

Cả 5 tháng **hai người đều đạt MỨC 3**. Mũi tên = trước → sau khi user cân tay.

---

## 14. Việc còn dang dở

**Rà sản phẩm từ comment trên Base.vn.** User muốn đếm số ảnh thật trong link
Google Drive ở phần comment của từng task design, để tính giờ cho sát thay vì đoán
từ tên task.

Tình trạng: mới xác nhận được **link Drive trong cột Mô tả/Feedback của file export
chỉ là brief đầu vào** (moodboard, reference), không phải sản phẩm — 39/117 task
tháng 9 có loại link này. Link sản phẩm nằm ở **comment**, mà comment **không có
trong file export** → phải mở từng task trên `workflow.base.vn` bằng trình duyệt
đã đăng nhập của user.

Chưa làm vì khối lượng lớn (tháng 9 có 77 task DESIGN MKT → ~154 lượt mở trang).
Nếu làm tiếp: **thử trên 1–2 task trước**, báo user chi phí/thời gian thực tế rồi
mới chạy cả loạt. Chỉ ĐỌC, tuyệt đối không comment hay sửa gì trên Base.

---

## 15. Tóm tắt cho lần chạy đầu tiên

```
1. Hỏi user: thư mục tháng ở đâu? (mặc định E:\Báo cáo hiệu suất\2026\<MMDD>\)
2. Liệt kê thư mục, xác nhận đủ file export + file báo cáo tháng trước.
3. Hỏi: Công chuẩn? TIMEOFF từng người? OT từng người?
4. Chạy pipeline (app hoặc script), build file.
5. Chạy đủ bộ kiểm tra §10.
6. Báo cáo: bảng điểm + bậc KPI, funnel lọc dữ liệu, phân tích BU,
   và NÓI RÕ chỗ nào là giả định / chỗ nào user cần tự quyết.
```
