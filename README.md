# My-Tools_VBA-Excel-for-work
Công cụ quản lý dữ liệu cho công ty Xuân Thảo
-----
## 📄 Mô tả

Đây là một class module VBA dùng để **đếm số lượng dòng** trong bảng dữ liệu `LoTrinh_Tong` trên sheet `TONG_HOP` của file Excel. Hàm `CountRecordFromTable` thực hiện việc này bằng cách:

* Nhận **biển số xe**, **ngày bắt đầu** và **ngày kết thúc** từ người dùng thông qua hộp thoại `InputBox`.
* Duyệt qua cột `BienSoXe` trong bảng.
* Kiểm tra xem mỗi dòng có thỏa điều kiện:

  * Biển số xe trùng khớp.
  * Ngày tương ứng nằm trong khoảng chỉ định.
* Trả về tổng số dòng thỏa mãn điều kiện.

## 🔧 Yêu cầu

* Một bảng có tên `LoTrinh_Tong` trên sheet `TONG_HOP`.
* Bảng cần có hai cột:

  * `BienSoXe`
  * `Ngay` (dữ liệu ngày tháng)

## 🚀 Cách sử dụng

1. Mở Excel chứa bảng dữ liệu.
2. Chạy macro `CountRecordFromTable`.
3. Nhập biển số xe, ngày bắt đầu và ngày kết thúc khi được yêu cầu.
4. Kết quả là số dòng thỏa điều kiện sẽ được trả về.
