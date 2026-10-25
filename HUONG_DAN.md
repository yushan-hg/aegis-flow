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

## 3b. Bốn tab trong app
Bản điện thoại chỉ để xem dữ liệu, không in, không chia đơn.
- **Trang chủ:** Nạp file · Bình khí (CO2F7, CO2F5, N2F13) · Xe không giao hàng · Dữ liệu đổi bản.
- **Việc làm:** hai tab con. **HSD**: các xe hôm nay có hàng phải nhặt theo hạn sử dụng (bấm số xe để lọc). **PCHome**: Hàng · SL · HSD · Bản/Lẻ, xếp số lượng từ lớn đến nhỏ.
- **Tra cứu:** **Dòng hàng** (Khách · Xe · Đơn · Mã · Hàng · SL) và **Tìm hàng** (Hàng · Xe · Đơn · SL).
- **Thêm:** những việc ngoài luồng: xem dữ liệu và cấu hình đang dùng (nạp tay hay có sẵn), dùng lại cấu hình có sẵn, ngôn ngữ (Việt · 中文 · English), giao diện sáng/tối, xoá dữ liệu trên máy (bấm hai lần).

Mọi danh sách là bảng phẳng: mỗi dòng một mặt hàng, có tiêu đề cột.

## 4. Khi sửa file cấu hình hoặc app
1. Sửa `vi_tri_kho_config.xlsx` / `doi_ban_config.xlsx` trong thư mục `Config` của app Windows như bình thường.
2. Chạy `tools\build_site.ps1` (tôi sẽ chạy giúp) — nó đọc lại hai file đó vào `docs/config.json`.
3. Trong GitHub Desktop: **Commit to main**, rồi **Push origin**. Điện thoại có mạng thì tự lấy cấu hình mới.

## 5. Thử ngoại tuyến
Sau khi cài xong, bật **chế độ máy bay** rồi mở app từ icon. Mở được và vẫn thấy dữ liệu đã nạp là ổn.
