# Aegis Flow trên iPhone

## 1. Đưa lên mạng bằng GitHub Pages (làm một lần)
1. Mở **GitHub Desktop**, đăng nhập. **File → Add local repository…** chọn thư mục `Mobile` này.
2. Bấm **Publish repository**, đặt tên `aegis-flow`, **bỏ dấu tích "Keep this code private"** (Pages miễn phí chỉ chạy với kho công khai).
3. Trên github.com vào kho → **Settings → Pages** → Deploy from a branch → nhánh `main`, thư mục `/docs` → Save.
4. Sau 1–2 phút app có địa chỉ `https://TEN-GITHUB.github.io/aegis-flow/`.

> Kho công khai: ai biết địa chỉ đều mở được app và xem được danh sách hàng, vị trí kho, danh sách khách trong `docs/config.json`. Dữ liệu giao hàng hằng ngày thì không lên mạng.

## 2. Cài lên iPhone
Mở địa chỉ bằng **Safari** → Chia sẻ → **Thêm vào Màn hình chính**. Lần mở đầu phải có mạng.

## 3. Dùng hằng ngày
1. Chuyển file Excel dữ liệu ngày (全虹…xlsx) sang iPhone (Zalo, AirDrop, iCloud), lưu vào **Files**.
2. Mở app → **Nạp file** → chọn file. **Không cần nạp file cấu hình**: `vi_tri_kho_config` và `doi_ban_config` đã có sẵn trong app.
3. Muốn dùng file cấu hình khác tạm thời thì cứ nạp tay, app sẽ ưu tiên file đó. Vào **Cài đặt → Dùng lại cấu hình có sẵn** để quay về.

## 3b. Năm tab trong app
- **Hôm nay:** số lượng tổng, số lượng nhặt lẻ (tải chia đơn), tiến độ HSD / Chia đơn / Đổi bản / PCHome, cảnh báo.
- **Chia đơn:** thẻ nhân viên (% tải), bấm xe → chọn người nhận hoặc xem chi tiết (theo đơn / bảng phẳng). Chia tự động, Cân bằng, Hoàn tác. Kết quả chia **không được lưu**: tắt app là mất.
- **Việc làm:** HSD đầu ca, Đổi bản, PCHome (bấm mặt hàng để nhập 整板 / 底 × 高), Đơn xe. Bấm dòng để đánh dấu ✓.
- **Tra cứu:** Dòng hàng, Tìm hàng, Theo xe.
- **Thêm:** nhân sự đi làm hôm nay (bật tắt, đổi năng lực), xe không giao hàng, dữ liệu & cài đặt, ngôn ngữ (Việt · 中文 · English), giao diện sáng/tối.

Bản điện thoại chỉ để xem, không có chức năng in. Danh sách nhân sự và xe lấy từ thư mục `Config` khi chạy `build_site.ps1`.

## 4. Khi sửa file cấu hình hoặc app
1. Sửa `vi_tri_kho_config.xlsx` / `doi_ban_config.xlsx` trong thư mục `Config` của app Windows như bình thường.
2. Chạy `tools\build_site.ps1` (tôi sẽ chạy giúp) — nó đọc lại hai file đó vào `docs/config.json`.
3. Trong GitHub Desktop: **Commit to main**, rồi **Push origin**. Điện thoại có mạng thì tự lấy cấu hình mới.

## 5. Thử ngoại tuyến
Sau khi cài xong, bật **chế độ máy bay** rồi mở app từ icon. Mở được và vẫn thấy dữ liệu đã nạp là ổn.
