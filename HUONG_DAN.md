# Nhập hàng loạt bài đăng — DHB Reup Facebook

Công cụ này giúp bạn soạn **nhiều bài đăng trong Excel** rồi nạp một lần vào tool DHB Reup,
và **sửa hàng loạt các dự án đã có**, thay vì gõ tay từng bài trong giao diện.

> **Quan trọng:** Công cụ nằm ở thư mục riêng `_NhapHangLoat`, **KHÔNG sửa gì trong file `.exe`/`.dll`
> của tool.** Nó chỉ ghi vào các file dự án (`.dhb`) trong thư mục `Data/Project` của tool — y hệt như
> khi bạn tự thêm/sửa bài rồi lưu dự án. Vì vậy khi nhà phát hành cập nhật tool, mọi thứ vẫn chạy bình thường.

---

## Chuyển sang máy khác dùng (4 bước)

1. **Cài Python 3** trên máy mới: tải ở https://www.python.org/downloads/ — khi cài **nhớ tick "Add Python to PATH"**.
2. **Copy cả thư mục `_NhapHangLoat`** sang máy mới, **đặt CẠNH thư mục "DHB Reup Facebook"** (chung một thư mục cha, giống như trên máy này). → Công cụ sẽ **tự tìm** thư mục dữ liệu của tool, không cần sửa code.
3. Nhấp đúp **`CaiThuVien.bat`** một lần (tự cài `openpyxl` + `customtkinter`).
4. Nhấp đúp **`Chay_GUI.bat`** để dùng.

**Nếu báo "không tìm thấy Data":** tạo file **`duong_dan_tool.txt`** trong `_NhapHangLoat`, dán vào **đường dẫn thư mục `Data`** của DHB Reup trên máy đó (vd `C:\Users\Tên\...\DHB Reup Facebook\DHB Reup Facebook\Data`), lưu lại rồi chạy lại.

**Lưu ý khi qua máy mới:**
- Danh sách trang/tài khoản lấy **trực tiếp** từ DHB Reup của máy mới → **xóa file `mau_bai_dang.xlsx`** (hoặc bấm *Cập nhật danh sách*) để dropdown khớp với các trang trên máy đó.
- Không cần mang theo thư mục `backups/` và các file `sua_*.xlsx` (là dữ liệu riêng của máy cũ).

---

## Cách dùng nhanh nhất: bản GIAO DIỆN (khuyên dùng)

Nhấp đúp **`Chay_GUI.bat`** → mở cửa sổ có nút bấm:
- **Mục 1 — Tạo dự án mới:** [Mở file Excel mẫu để điền] rồi [Tạo / ghi vào dự án].
- **Mục 2 — Sửa dự án có sẵn:** chọn dự án ở ô xổ xuống → [Xuất ra Excel để sửa]; sửa xong chọn file → [Nạp lại].
- **Mục 3 — Tiện ích:** [Cập nhật danh sách trang/tài khoản], [Mở thư mục].
- Góc trên phải báo **tool đang mở hay đã tắt**; khu **Nhật ký** ở dưới hiện kết quả từng bước.

Mọi cảnh báo (trùng tên dự án, chọn ghi đè/nối thêm, xác nhận…) đều hiện bằng **hộp thoại bấm nút**, không phải gõ phím.

> Bản dòng lệnh (`Chay_NhapHangLoat.bat`) bên dưới vẫn dùng được như cũ — cùng chung logic, chọn bản nào tuỳ bạn.

---

## (Tùy chọn) Bản dòng lệnh — Menu khi chạy

Nhấp đúp **`Chay_NhapHangLoat.bat`** sẽ hiện menu:

```
  1. Tạo dự án mới từ Excel (mau_bai_dang.xlsx)
  2. Xuất 1 dự án ra Excel để SỬA
  3. Nạp Excel đã sửa  ->  ghi đè lại dự án
  4. Cập nhật danh sách trang/tài khoản trong file Excel
  0. Thoát
```

> **Về dropdown (ô xổ xuống):** danh sách trang/tài khoản trong dropdown là *ảnh chụp lúc tạo file*, **không tự cập nhật**.
> Khi bạn thêm/xóa fanpage hay tài khoản trong tool, chọn **mục 4** để kéo lại danh sách mới nhất vào các file Excel
> (giữ nguyên dữ liệu đã điền). Kể cả chưa cập nhật, lúc **nạp** vẫn khớp theo dữ liệu sống của tool — gõ tay tên
> trang mới vẫn đúng; dropdown chỉ để đỡ gõ sai.

> Luôn **TẮT hẳn tool DHB Reup trước** khi chọn 1 hoặc 3 (ghi dữ liệu). Nếu tool đang mở, công cụ sẽ báo và dừng.
> Mọi lần ghi đều **tự backup** dự án cũ vào thư mục `backups/`.

---

## A) Tạo dự án mới (mục 1)

### Bước 1 — Tạo file Excel mẫu
Chạy công cụ, chọn **1**. Lần đầu chưa có file, nó tự tạo **`mau_bai_dang.xlsx`** rồi thoát.

### Bước 2 — Điền bài vào Excel
Mở `mau_bai_dang.xlsx`. Có 3 sheet:
- **BaiDang** — mỗi dòng là 1 bài. Các cột:

  | Cột | Bắt buộc | Ý nghĩa |
  |---|---|---|
  | **Dự án** | | Tên dự án cho bài đó. **Nhiều bài cùng tên → gom vào 1 dự án.** Để trống → vào dự án mặc định (công cụ sẽ hỏi tên). ⇒ **1 file Excel tạo được nhiều dự án cùng lúc.** |
  | **Nguồn (URL)** | ✔ | Link file trực tiếp (Direct Link, vd `https://pub-...r2.dev/...mp4`), link YouTube/Facebook, hoặc đường dẫn file trên máy |
  | **Loại nguồn** | | `Direct Link` / `Youtube` / `Facebook` / `File`… **Để trống = tự nhận dạng** (link http lạ → Direct Link, link youtube → Youtube, đường dẫn `C:\…` → File) |
  | **Caption / Mô tả** | | Nội dung bài (xuống dòng thoải mái trong ô) |
  | **Trang đăng** | ✔ | Tên **trang** sẽ đăng, HOẶC một **nhóm fanpage**: chọn `GROUP: <tên nhóm>` (vd `GROUP: lamdep`) — đăng vào **tất cả page trong nhóm** đó. Chọn từ dropdown, hoặc copy từ sheet `DanhSachTrang` (dòng có Loại = GROUP). |
  | **Tài khoản** | | Chỉ cần khi một tên trang trùng ở 2 tài khoản |
  | **Đăng dạng** | | `Reels` (mặc định), `Bài viết`, hoặc `Story` |
  | **Lịch đăng** | | `2026-07-05 08:30`. Để trống = đăng luôn |
  | **Nhóm tùy chọn** | | Gắn preset (link Shopee…) — copy tên từ sheet `NhomTuyChon` |
  | **Tiêu đề** | | Tùy chọn |
  | **Link** | | Tùy chọn |

- **DanhSachTrang** — danh sách tất cả trang + tài khoản + **nhóm** tool đang có (để copy đúng tên).
- **DanhSachNhom** — mỗi **nhóm fanpage** (`GROUP: …`) gồm những page nào (đăng vào nhóm = đăng lên tất cả page đó).
- **NhomTuyChon** — danh sách các nhóm preset (link Shopee…) hiện có.

> 💡 Các cột **Loại nguồn, Trang đăng, Tài khoản, Đăng dạng** (và **Trạng thái** trong file sửa) có **ô xổ xuống (dropdown)** —
> bấm vào ô là chọn được, khỏi gõ tay cho khỏi sai. (Có 1 sheet ẩn tên `Lists` chứa danh sách này — đừng xóa.)
> Trang bị trùng tên ở 2 tài khoản thì chọn thêm ô **Tài khoản** để rõ.

Điền xong **lưu file** (vẫn tên `mau_bai_dang.xlsx`).

### Bước 3 — Nạp vào tool
1. **TẮT hẳn tool DHB Reup**.
2. Chạy công cụ → chọn **1** → gõ **tên dự án mặc định** (Enter = `NhapHangLoat`).
   Các bài có ghi cột **"Dự án"** sẽ vào đúng dự án tên đó; bài để trống → vào dự án mặc định vừa gõ.
3. Nó báo đã tạo mấy dự án, mỗi dự án mấy bài.
4. Mở tool → **Load từng dự án** → kiểm tra → bấm chạy.

> **Muốn nhiều dự án khác nhau?** Chỉ cần điền tên khác nhau ở cột **"Dự án"** — mỗi tên thành 1 file `.dhb` riêng.
> (Hoặc chạy mục 1 nhiều lần, mỗi lần gõ tên dự án khác — nhưng dùng cột "Dự án" thì làm 1 lần được hết.)

> **Nếu gõ trùng tên một dự án ĐÃ CÓ?** Công cụ sẽ **cảnh báo** và cho bạn chọn:
> - **Ghi đè (g)** — xóa hết bài cũ của dự án đó, thay bằng bài trong Excel.
> - **Nối thêm (n)** — giữ nguyên bài cũ, thêm bài mới vào cuối.
> - Phím khác — huỷ, không ghi gì.
>
> Dù chọn gì, bản cũ **luôn được backup** vào `backups/<tên>_<ngày giờ>.dhb` trước — khôi phục được nếu lỡ tay.

---

## B) Sửa hàng loạt một dự án đã có (mục 2 → sửa Excel → mục 3)

Dùng khi bạn muốn sửa nhiều bài trong một dự án sẵn có (đổi caption, nguồn, trang, lịch…) rồi chạy lại.

1. **TẮT tool** → chạy công cụ → chọn **2 (Xuất)** → chọn dự án cần sửa.
   Nó tạo file **`sua_<tên dự án>.xlsx`** chứa **toàn bộ bài hiện có** của dự án đó.
2. Mở file `sua_...xlsx`, sửa ở sheet **BaiDang**. Lưu ý:
   - **Cột cuối cùng đang bị ẩn** là *bản gốc từng bài* — **KHÔNG sửa/xóa cột này**. Nhờ nó, mọi trường bạn
     không đụng tới đều được **giữ nguyên** (sửa không mất dữ liệu).
   - Cột **"Trạng thái"**: muốn 1 bài **chạy lại** thì đặt = `Not start`. Để nguyên `Hoàn thành` thì bài đó
     **không chạy lại** (tránh đăng trùng) — dù bạn có sửa caption đi nữa.
   - Xóa dòng = bỏ bài đó khỏi dự án. Thêm dòng mới = thêm bài mới (để trống cột ẩn).
3. Lưu file → chạy công cụ → chọn **3 (Nạp)** → chọn đúng file `sua_...` → gõ `y` xác nhận.
   Nó **ghi đè lại chính dự án đó** (đã backup bản cũ) và báo có bao nhiêu bài sẽ chạy.
4. Mở tool → load lại dự án → bấm chạy.

---

## An toàn
- **Tự động backup:** mỗi lần nạp, file dự án cũ được copy vào thư mục `backups/` kèm ngày giờ.
- **Không đụng dự án khác:** công cụ chỉ ghi vào `NhapHangLoat.dhb`. Các dự án cũ của bạn không bị ảnh hưởng.
- **Excel là nguồn duy nhất:** mỗi lần chạy, `NhapHangLoat.dhb` được dựng lại đúng theo Excel
  (chạy lại không bị nhân đôi bài). Muốn *nối thêm* thay vì dựng lại: chạy với `--noi-them`.

## Nguồn đa dạng — công cụ tự nhận loại & khớp đúng định dạng tool

Cột **"Loại nguồn"** để trống thì công cụ **tự nhận dạng** từ URL/đường dẫn. Cũng có thể ghi rõ.

| Loại nguồn | Điền vào cột "Nguồn" | Trạng thái |
|---|---|---|
| **Direct Link** | link file trực tiếp `https://…mp4` (vd R2/CDN) | ✅ khớp GUI (Source rỗng, tải từ Url) |
| **File** | đường dẫn file có đuôi video `C:\…\video.mp4` | ✅ khớp bài File thật của bạn |
| **Youtube** | `https://youtube.com/watch?v=…` hoặc `youtu.be/…` | ✅ tool tự tải (yt-dlp) |
| **Facebook** | `https://facebook.com/reel/…` | ✅ dạng đúng (bài reup FB cần tool tự resolve khi chạy) |
| **Folder** | đường dẫn thư mục `C:\…\video_folder` (không đuôi) | ⚙️ nhận dạng được nhưng **chưa kiểm chứng** |
| Youtube-DL / Dailymotion / List Text / Keyword / Status / AI Research / Driver | tùy loại | ⚙️ hỗ trợ khi có mẫu |

**Cơ chế "tự học khuôn":** công cụ lấy khuôn cho mỗi loại nguồn **từ chính các bài cùng loại bạn từng tạo** (trong các dự án `.dhb` cũ), nên định dạng luôn khớp phiên bản tool của bạn. Loại nào bạn **chưa từng dùng**, công cụ sẽ báo "chưa có khuôn". Cách xử lý: **tạo 1 bài loại đó trong tool → lưu dự án → chạy lại công cụ** — nó sẽ tự nhặt khuôn đó và từ đó nhập hàng loạt loại này chuẩn 100%.

> Với loại mới, luôn chạy thử **1 bài** trước khi nạp số lượng lớn.

## Sau này muốn dùng Google Sheets
Được. Chỉ cần đổi phần "đọc dữ liệu vào" (mình đã tách riêng), phần ghi vào tool giữ nguyên.
Khi cần, nhắn mình bật chế độ đọc thẳng từ Google Sheets của bạn.

## Nếu chuyển tool sang máy khác / đổi đường dẫn
Mở `import_bai_dang.py`, sửa dòng `TOOL_DATA = ...` ở đầu file cho trỏ đúng thư mục `Data` của tool.
