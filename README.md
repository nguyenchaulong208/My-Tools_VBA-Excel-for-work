# My-Tools_VBA-Excel-for-work
Công cụ quản lý dữ liệu cho công ty Xuân Thảo
------

# 📊 VBA Route Management Tool

Dự án này là một tập hợp các macro VBA dùng để **quản lý và xử lý dữ liệu lộ trình xe** trong Excel. Bộ công cụ bao gồm các chức năng như đếm số bản ghi, lọc dữ liệu theo biển số và thời gian, tính giờ làm thêm, và xuất dữ liệu ra bảng tổng hợp.

---

## 📁 Cấu trúc mã nguồn

| Tên file                 | Mô tả                                                                                                           |
| ------------------------ | --------------------------------------------------------------------------------------------------------------- |
| `DataLoTrinh.cls`        | Lớp chứa thông tin tìm kiếm: biển số xe, ngày bắt đầu, ngày kết thúc, và tên sheet.                             |
| `ThongTinLoTrinh.cls`    | Lớp chứa chi tiết lộ trình: ngày, địa điểm, giờ bắt đầu/kết thúc, số km, tài xế, tuyến đường, số lượng vé,...   |
| `CountRecord.bas`        | Hàm chính `CountRecordFromTable` dùng để đếm số dòng dữ liệu thỏa mãn điều kiện biển số xe và khoảng thời gian. |
| `AddRow.bas`             | Thêm dòng mới vào vùng dữ liệu tên là `data_Export` nếu số dòng hiện tại chưa đủ.                               |
| `GetRecord.bas`          | Lọc và thu thập các dòng dữ liệu phù hợp để đưa vào bộ sưu tập `ThongTinLoTrinh`.                               |
| `OverTime.bas`           | Tính số phút làm thêm dựa trên giờ thực tế và giờ chuẩn theo biển số xe.                                        |
| `WriteData.bas`          | Ghi dữ liệu từ `ThongTinLoTrinh` ra vùng tên (named ranges) trong sheet `Export_LoTrinh`.                       |
| `optimized_vba_code.bas` | Tập hợp mã VBA đã được tối ưu, bao gồm định nghĩa các lớp và thủ tục xử lý chính.                               |

---

## 🧰 Tính năng chính

* ✅ Nhập điều kiện từ người dùng (biển số xe, ngày bắt đầu/kết thúc).
* 🔎 Lọc dữ liệu theo điều kiện.
* 🧮 Tính số bản ghi phù hợp.
* 🕒 Tính giờ làm thêm.
* 🧾 Ghi dữ liệu ra bảng tổng hợp.
* ➕ Tự động thêm dòng nếu số dòng chưa đủ.

---

## 🚀 Hướng dẫn sử dụng

1. Mở file Excel chứa dữ liệu gốc (sheet `TONG_HOP`, bảng `LoTrinh_Tong`).
2. Nhấn Alt + F11 để mở trình soạn thảo VBA.
3. Chạy macro `AddRowNameRange` để thêm dòng và xử lý toàn bộ dữ liệu.
4. Nhập thông tin khi được yêu cầu.
5. Dữ liệu đã xử lý sẽ được xuất ra sheet `Export_LoTrinh`.

---

## 📌 Yêu cầu

* Excel có bảng tên `LoTrinh_Tong` và `ThongTinChung`.
* Các tên vùng (named ranges) trong sheet `Export_LoTrinh` phải được định nghĩa:

  * `Ngay_Ex`, `TaiXe_Ex`, `DiaDiem_Ex`, `StartTime_Ex`, `EndTime_Ex`, `OverTime_Ex`, `KM_Ex`, `VeVETC_Ex`, `SoLuong_Ex`



