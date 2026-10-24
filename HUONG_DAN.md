# Aegis Flow trên iPhone — hướng dẫn đưa lên và dùng

## 1. Đưa lên mạng (làm một lần, mất khoảng 5 phút)
iPhone chỉ cho chạy ngoại tuyến khi trang có HTTPS, nên cần đặt thư mục `site` ở một nơi có địa chỉ riêng. Cách dễ nhất là **Netlify** (miễn phí):

1. Vào https://app.netlify.com/drop và đăng ký tài khoản (có thể dùng email hoặc Google).
2. Kéo cả thư mục **`Mobile\site`** thả vào khung trên trang.
3. Netlify cho một địa chỉ dạng `https://tên-ngẫu-nhiên.netlify.app`. Có thể đổi tên trong Site settings → Change site name.

Trong thư mục `site` chỉ có file của app — **không có dữ liệu hay file cấu hình của bạn**, nên để công khai cũng không lộ gì.

## 2. Cài lên iPhone
1. Mở địa chỉ trên bằng **Safari** (bắt buộc Safari, không phải Chrome).
2. Bấm nút **Chia sẻ** → **Thêm vào Màn hình chính**.
3. Mở app từ icon vừa tạo. Lần mở đầu phải có mạng để máy tải sẵn app.

## 3. Dùng hằng ngày
1. Chuyển file Excel dữ liệu ngày (全虹…xlsx) sang iPhone (Zalo, AirDrop hoặc iCloud), lưu vào ứng dụng **Files**.
2. Mở app → **Nạp file** → chọn file. Chọn được nhiều file một lần.
3. Hai file cấu hình `vi_tri_kho_config.xlsx` và `doi_ban_config.xlsx` chỉ cần nạp lần đầu và khi bạn sửa chúng.

## 4. Thử ngoại tuyến
Sau khi cài xong, bật **chế độ máy bay** rồi mở app từ icon. Nếu mở được và vẫn thấy dữ liệu đã nạp thì phần ngoại tuyến đã chạy. Nếu không mở được thì báo lại để sửa.

## 5. Khi app được cập nhật
Chạy lại `tools\build_site.ps1` rồi kéo thả thư mục `site` mới vào Netlify (Deploys → kéo vào ô deploy). Lần mở app kế tiếp có mạng, điện thoại tự tải bản mới.
