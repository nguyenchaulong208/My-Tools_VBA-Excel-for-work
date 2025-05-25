
# 📊 VBA Route Management Tool

## Tổng quan
Dự án này là một tập hợp các macro VBA dùng để **quản lý và xử lý dữ liệu lộ trình xe** trong Excel. Bộ công cụ bao gồm các chức năng như đếm số bản ghi, lọc dữ liệu theo biển số và thời gian, tính giờ làm thêm, và xuất dữ liệu ra bảng tổng hợp.

## Tính năng
- **Trích xuất dữ liệu lộ trình**: Lấy dữ liệu lộ trình từ bảng chính (`LoTrinh_Tong`) dựa trên thông tin đầu vào như biển số xe và khoảng thời gian.
- **Quản lý dòng động**: Tự động thêm hoặc xóa các dòng trong bảng tính đầu ra (`Export_LoTrinh`) để khớp với số lượng bản ghi được trích xuất.
- **Tính toán thời gian làm thêm giờ**: Tính thời gian làm thêm giờ dựa trên khung giờ làm việc tiêu chuẩn của từng phương tiện.
- **Tính toán doanh thu**: Tính tổng doanh thu, bao gồm cước tháng, phí làm thêm giờ và các khoản phí bổ sung như vé VETC, có tính đến thuế.
- **Kiểm tra dữ liệu**: Đánh dấu các ô trống trong các cột quan trọng (ví dụ: `SoKmDaSuDung`) để đảm bảo chất lượng dữ liệu.
- **Tích hợp với Excel**: Sử dụng bảng Excel (`ListObjects`) và các vùng được đặt tên để thao tác dữ liệu hiệu quả.

## Cấu trúc dự án
Dự án bao gồm các mô-đun VBA và mô-đun lớp, mỗi mô-đun đảm nhiệm một vai trò cụ thể:

- **MainModule.bas**: Chứa thủ tục chính (`Main`) điều phối quy trình trích xuất và xử lý dữ liệu.
- **GetRecord.bas**: Trích xuất dữ liệu lộ trình từ bảng `LoTrinh_Tong` và lưu vào một bộ sưu tập để xử lý.
- **WriteData.bas**: Ghi dữ liệu lộ trình đã xử lý vào bảng tính `Export_LoTrinh`.
- **CountRecord.bas**: Đếm số bản ghi khớp với tiêu chí do người dùng chỉ định (biển số xe và khoảng thời gian).
- **OverTime.bas**: Tính thời gian làm thêm giờ dựa trên thời gian bắt đầu và kết thúc so sánh với giờ làm việc tiêu chuẩn.
- **Calculate.bas**: Thực hiện tính toán doanh thu, bao gồm cước tháng, làm thêm giờ và thuế.
- **CheckCell.bas**: Kiểm tra dữ liệu bằng cách đánh dấu các ô trống trong cột `SoKmDaSuDung`.
- **ThongTinLoTrinh.cls**: Mô-đun lớp định nghĩa cấu trúc cho dữ liệu lộ trình (ví dụ: ngày, địa điểm, quãng đường, tài xế).
- **DataLoTrinh.cls**: Mô-đun lớp lưu trữ các tham số đầu vào như biển số xe, khoảng thời gian và tên bảng tính.

## Yêu cầu
- **Microsoft Excel**: Yêu cầu Excel có bật VBA (macro phải được bật).
- **Thiết lập Workbook**:
  - **Bảng tính `TONG_HOP`**:
    - Chứa bảng `LoTrinh_Tong` với các cột:
      - `BienSoXe`: Biển số xe.
      - `Ngay`: Ngày lộ trình.
      - `DiaDiem`: Địa điểm.
      - `ThoiGianBatDau`: Thời gian bắt đầu.
      - `ThoiGianKetThuc`: Thời gian kết thúc.
      - `SoKmBatDau`: Số km bắt đầu.
      - `SoKmKetThuc`: Số km kết thúc.
      - `SoKmDaSuDung`: Quãng đường đã sử dụng.
      - `TongTienVetc`: Tổng tiền vé VETC.
      - `SoLuongVe`: Số lượng vé.
      - `TaiXe`: Tên tài xế.
      - `TuyenDuong`: Tuyến đường.
      - `CongTy`: Khách hàng/công ty.
  - **Bảng tính `THONG_TIN_CHUNG`**:
    - Chứa bảng `ThongTinChung` với các cột:
      - `BienSoXe`: Biển số xe.
      - `BatDau`: Giờ làm việc bắt đầu tiêu chuẩn.
      - `KetThuc`: Giờ làm việc kết thúc tiêu chuẩn.
      - `DoanhThuThang`: Doanh thu tháng cố định.
      - `DonGiaNgayChuNhat`: Đơn giá ngày Chủ Nhật.
      - `DonGiaKmVuot`: Đơn giá km vượt.
      - `DonGiaQuaGio`: Đơn giá làm thêm giờ.
  - **Bảng tính `Export_LoTrinh`**:
    - Chứa các vùng được đặt tên (Named Ranges):
      - `data_Export`: Vùng dữ liệu chính của bảng lộ trình.
      - `Ngay_Ex`: Cột ngày.
      - `TaiXe_Ex`: Cột tài xế.
      - `DiaDiem_Ex`: Cột địa điểm.
      - `StartTime_Ex`: Cột thời gian bắt đầu.
      - `EndTime_Ex`: Cột thời gian kết thúc.
      - `OverTime_Ex`: Cột thời gian làm thêm giờ.
      - `KM_Ex`: Cột quãng đường.
      - `VeVETC_Ex`: Cột tổng tiền vé VETC.
      - `SoLuong_Ex`: Cột số lượng vé.
      - `SumOverTime_Ex`: Ô tổng thời gian làm thêm giờ.
      - `SumKM_Ex`: Ô tổng quãng đường.
      - `TT_TongThanhTien_Ex`: Ô tổng doanh thu.
      - `TT_TienThue_Ex`: Ô tổng tiền thuế.
      - `TT_TongCong_Ex`: Ô tổng cộng (doanh thu + thuế).

## Cài đặt
1. **Tải hoặc sao chép**: Tải kho lưu trữ hoặc sao chép các tệp VBA vào máy cục bộ.
2. **Nhập tệp VBA**:
   - Mở workbook Excel.
   - Nhấn `Alt + F11` để mở trình chỉnh sửa VBA.
   - Nhấp chuột phải vào dự án trong Project Explorer, chọn `Import File`, nhập tất cả tệp `.bas` và `.cls`.
3. **Thiết lập Workbook**:
   - Tạo các bảng tính `TONG_HOP`, `THONG_TIN_CHUNG`, `Export_LoTrinh`.
   - Tạo bảng `LoTrinh_Tong` và `ThongTinChung` với các cột như mô tả.
   - Định nghĩa các vùng được đặt tên (Named Ranges) trong Excel khớp với mã (ví dụ: `data_Export`, `Ngay_Ex`).
4. **Bật Macro**: Đảm bảo macro được bật trong Excel.

## Hướng dẫn sử dụng
1. **Chạy thủ tục chính**:
   - Mở workbook Excel.
   - Nhấn `Alt + F8`, chọn `Main`, nhấp `Run`.
   - Nhập biển số xe, ngày bắt đầu, ngày kết thúc qua hộp thoại nhập liệu.
2. **Kết quả đầu ra**:
   - Bảng `Export_LoTrinh` được cập nhật với dữ liệu lộ trình (ngày, tài xế, quãng đường, v.v.).
   - Các ô tổng hợp (`SumOverTime_Ex`, `SumKM_Ex`, `TT_TongThanhTien_Ex`, v.v.) được điền giá trị.
   - Hộp thoại thông báo xác nhận hoàn tất, cho biết nếu có dòng thừa bị xóa.
3. **Kiểm tra dữ liệu**:
   - Chạy macro `CheckCellEmpty` để đánh dấu ô trống trong cột `SoKmDaSuDung`.

## Ví dụ quy trình
1. Người dùng chạy macro `Main`.
2. Đầu vào: Biển số xe (`29A-12345`), ngày bắt đầu (`01/05/2025`), ngày kết thúc (`31/05/2025`).
3. Mã thực hiện:
   - Đếm bản ghi khớp trong `LoTrinh_Tong`.
   - Điều chỉnh dòng trong `Export_LoTrinh`.
   - Trích xuất và ghi dữ liệu vào `Export_LoTrinh`.
   - Tính thời gian làm thêm giờ, doanh thu, cập nhật ô tổng hợp.
4. Kết quả: Bảng `Export_LoTrinh` chứa dữ liệu lộ trình và tóm tắt tài chính.

## Lưu ý
- **Hiệu suất**: Mã tắt cập nhật màn hình và tính toán tự động để tăng tốc. Cân nhắc thêm `Application.ScreenUpdating = True` và `Application.Calculation = xlCalculationAutomatic` vào cuối `Main` nếu cần.
- **Xử lý lỗi**: Mã giả định đầu vào hợp lệ. Cân nhắc thêm xử lý lỗi cho ngày không hợp lệ, bảng thiếu hoặc đầu vào trống.
- **Tính toàn vẹn dữ liệu**: Đảm bảo bảng `LoTrinh_Tong` và `ThongTinChung` được điền đúng để tránh lỗi.

