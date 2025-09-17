# CCTV-SN-Scan - Công cụ Quét và Lưu trữ Mã Vạch

Đây là một ứng dụng web đơn giản được xây dựng bằng Python và Flask, cho phép người dùng tải lên hình ảnh chứa mã QR (QR code) hoặc mã vạch (barcode) để giải mã, lưu trữ thông tin và xuất dữ liệu ra tệp Excel.

Ứng dụng được thiết kế để xử lý thông tin sản phẩm cho ba thương hiệu cụ thể: **IMOU**, **HIKVISION**, và **DAHUA**.

## Tính năng chính

- **Quét mã từ ảnh**: Tải lên tệp ảnh (JPG, PNG,...) để tự động giải mã mã QR/barcode.
- **Quản lý dữ liệu theo thương hiệu**: Dữ liệu được phân loại và lưu trữ riêng biệt cho từng thương hiệu.
- **Lưu trữ linh hoạt**: Dữ liệu được lưu dưới dạng tệp JSON, dễ dàng đọc và xử lý.
- **Xuất báo cáo Excel**: Xuất danh sách dữ liệu đã quét ra tệp `.xlsx` được định dạng chuyên nghiệp (tiêu đề, màu sắc, viền, độ rộng cột tự động).
- **Sao lưu an toàn**: Tự động sao lưu tệp dữ liệu hiện tại trước khi thực hiện thao tác xóa.
- **Giao diện web đơn giản**: Dễ dàng thao tác và sử dụng ngay trên trình duyệt.

## Công nghệ sử dụng

- **Backend**:
  - [Python 3](https://www.python.org/)
  - [Flask](https://flask.palletsprojects.com/): Một micro web framework để xây dựng ứng dụng.
- **Thư viện xử lý**:
  - `pyzbar`: Để giải mã mã QR và barcode.
  - `opencv-python-headless`: Để xử lý hình ảnh đầu vào.
  - `pandas` & `openpyxl`: Để tạo và định dạng tệp Excel.
  - `numpy`: Để thực hiện các phép toán trên dữ liệu hình ảnh.
- **Frontend**:
  - HTML5
  - CSS3

## Cài đặt và Khởi chạy

### Yêu cầu

- Python 3.8+
- `pip` (trình quản lý gói của Python)

### Hướng dẫn cài đặt

1.  **Tải mã nguồn:**
    Tải và giải nén mã nguồn vào một thư mục trên máy tính của bạn.

2.  **Tạo và kích hoạt môi trường ảo:**
    Mở Command Prompt trong thư mục dự án và chạy các lệnh sau:

    ```bash
    # Tạo môi trường ảo
    python -m venv venv

    # Kích hoạt môi trường ảo
    .\venv\Scripts\activate
    ```

3.  **Cài đặt các thư viện cần thiết:**
    Sử dụng tệp `requirements.txt` để cài đặt tất cả các phụ thuộc:

    ```bash
    pip install -r requirements.txt
    ```

### Khởi chạy ứng dụng

Cách đơn giản nhất để chạy ứng dụng là thực thi tệp `run.bat`. Thao tác này sẽ tự động kích hoạt môi trường ảo và khởi động máy chủ.

Sau khi chạy, mở trình duyệt web và truy cập vào một trong các địa chỉ sau:
- `http://127.0.0.1:5005`
- `http://<DIA_CHI_IP_CUA_BAN>:5005`

## Hướng dẫn sử dụng

1.  Từ trang chủ, chọn một trong ba thương hiệu: **IMOU**, **HIK**, hoặc **DAHUA**.
2.  Trên trang của thương hiệu, nhấn vào nút **"Chọn tệp"** và tải lên một hình ảnh có chứa mã vạch hoặc mã QR.
3.  Dữ liệu được giải mã sẽ hiển thị trong phần **"Kết quả quét"**.
4.  Bạn có thể thêm ghi chú vào ô **"Note"** và nhấn **"Lưu kết quả"**.
5.  Dữ liệu mới sẽ được thêm vào bảng **"Danh sách dữ liệu đã lưu"**.
6.  Sử dụng các nút chức năng:
    - **"Export to Excel"**: Để tải về tệp Excel chứa toàn bộ dữ liệu trong bảng.
    - **"Clear Data"**: Để xóa toàn bộ dữ liệu (sau khi đã sao lưu).
