# Công Cụ Tự Động Sửa File PST/OST

Đây là một công cụ có giao diện đồ họa (GUI) được viết bằng Python để tự động hóa tiện ích `SCANPST.EXE` của Microsoft. Mục đích chính là giúp người dùng sửa chữa hàng loạt các file dữ liệu Outlook (.pst, .ost) một cách nhanh chóng và tiện lợi, thay vì phải thao tác thủ công cho từng file.

![Ảnh chụp màn hình của ứng dụng](httpsd://raw.githubusercontent.com/danhcp/scanpst-py/main/screenshot.png)

## Tính năng chính

- **Giao diện đồ họa thân thiện:** Dễ dàng sử dụng ngay cả với người không có chuyên môn kỹ thuật.
- **Tự động phát hiện Office:** Tự động tìm kiếm `SCANPST.EXE` từ nhiều phiên bản Office khác nhau (365, 2019, 2016, 2013, 2010).
- **Xử lý hàng loạt:** Cho phép chọn và sửa nhiều file `.pst`/`.ost` cùng lúc.
- **Tùy chọn Backup:** Cho phép người dùng quyết định có tạo file backup (`.bak`) trước khi sửa hay không.
- **Lặp lại quá trình sửa:** Tùy chọn lặp lại quá trình sửa chữa từ 1 đến 3 lần để đảm bảo khắc phục lỗi triệt để.
- **Theo dõi trực quan:** Hiển thị trạng thái chi tiết của quá trình sửa lỗi theo thời gian thực ngay trên giao diện.

## Yêu cầu

- **Hệ điều hành:** Windows 10.
- **Microsoft Office:** Cần phải được cài đặt trên máy vì công cụ này chỉ tự động hóa `SCANPST.EXE` có sẵn của Office.

## Cài đặt và Sử dụng

### Cách 1: Sử dụng file đã build (Khuyên dùng)

1.  Đi đến thư mục `dist`.
2.  Chạy file `scanpst.exe` và sử dụng. Không cần cài đặt gì thêm.

### Cách 2: Chạy từ mã nguồn (Dành cho lập trình viên)

1.  Đảm bảo bạn đã cài đặt Python 3.
2.  Cài đặt các thư viện cần thiết bằng pip:
    ```sh
    pip install ttkbootstrap pywinauto
    ```
3.  Chạy ứng dụng bằng lệnh:
    ```sh
    python scanpst.py
    ```

## Hướng dẫn sử dụng

1.  **Chạy ứng dụng:** Mở file `scanpst.exe` hoặc chạy script `scanpst.py`.
2.  **Chọn phiên bản Office:** Ứng dụng sẽ tự động chọn phiên bản Office tìm thấy. Nếu có nhiều phiên bản, bạn có thể chọn từ danh sách.
3.  **Chọn file:** Nhấn nút **"Chọn Files..."** để chọn một hoặc nhiều file `.pst` hoặc `.ost` bạn muốn sửa.
4.  **Tùy chọn Backup:** Mặc định, `SCANPST.EXE` sẽ tạo một file backup. Nếu bạn không muốn, hãy bỏ tick ở ô **"Tạo file backup (.bak)..."**.
5.  **Chọn số lần lặp:** Chọn số lần bạn muốn lặp lại toàn bộ quá trình sửa chữa (từ 1 đến 3).
6.  **Bắt đầu:** Nhấn nút **"BẮT ĐẦU SỬA LỖI"**.
7.  **Theo dõi:** Theo dõi tiến trình trong khung **"Trạng Thái Hoạt Động"**. Mọi bước sẽ được ghi lại ở đây.
8.  **Hoàn tất:** Khi quá trình kết thúc, ứng dụng sẽ hiển thị thông báo tổng kết.

## Build ứng dụng

Nếu bạn muốn tự build file `.exe` từ mã nguồn, hãy chạy file `build.cmd`. File thực thi sẽ được tạo trong thư mục `dist`.

## Tác giả

Công cụ được phát triển bởi **@danhcp**.
