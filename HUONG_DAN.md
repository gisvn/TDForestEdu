# Hướng dẫn tạo TDForestEdu.dmg cho macOS

**Lưu ý quan trọng:** file .dmg/.app cho Mac KHÔNG thể tạo từ Windows (PyInstaller không hỗ trợ build chéo).
Bắt buộc build trên máy Mac, hoặc dùng máy Mac miễn phí của GitHub Actions (Cách 2).

## Cách 1: Có máy Mac
1. Cài Python 3.11/3.12 từ python.org (khuyên dùng, đừng dùng Python có sẵn của Apple).
2. Copy cả thư mục này sang Mac, mở Terminal tại thư mục đó.
3. Chạy:
   ```
   chmod +x build_mac.sh
   ./build_mac.sh
   ```
4. Kết quả: `dist/TDForestEdu_1.0.0_mac_arm64.dmg` (máy M1/M2/M3) hoặc `..._x86_64.dmg` (máy Intel).
   Build trên máy chip nào thì chạy trên máy chip đó.

## Cách 2: Không có máy Mac (GitHub Actions, miễn phí)
1. Tạo repo trên GitHub, upload toàn bộ thư mục này (kể cả thư mục ẩn `.github`).
2. Vào tab **Actions** -> chọn **Build macOS DMG** -> **Run workflow**.
3. Đợi ~10 phút, tải 2 file .dmg (Apple Silicon + Intel) ở mục **Artifacts**.

## Khi người dùng mở app lần đầu
App chưa ký bằng Apple Developer ID nên macOS sẽ báo "không thể xác minh nhà phát triển". Cách mở:
- Chuột phải vào TDForestEdu.app -> **Open** -> **Open**, hoặc
- System Settings -> Privacy & Security -> **Open Anyway**, hoặc
- Terminal: `xattr -cr /Applications/TDForestEdu.app`

Muốn hết hẳn cảnh báo: cần tài khoản Apple Developer (99 USD/năm) để ký `Developer ID` + notarize.

## Ghi chú về code
- Font "Segoe UI" không có trên Mac: Qt tự thay font hệ thống, matplotlib có thể báo warning "findfont" rồi dùng DejaVu Sans. Không lỗi, chỉ khác font. Nếu muốn đẹp hơn trên Mac, thêm "Helvetica Neue" / "Arial" vào danh sách font ưu tiên.
- Cấu hình QSettings trên Mac được lưu ở `~/Library/Preferences/com.tdforest-software.TDForest_TracDoRung.plist`.
