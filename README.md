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
- **Microsoft Excel**: Dự án yêu cầu Excel có bật VBA (cần bật macro).
- **Thiết lập Workbook**:
  - Một bảng tính tên `TONG_HOP` chứa bảng `LoTrinh_Tong` với các cột như `BienSoXe`, `Ngay`, `DiaDiem`, v.v.
  - Một bảng tính tên `THONG_TIN_CHUNG` chứa bảng `ThongTinChung` cho các cài đặt cụ thể của phương tiện (ví dụ: giờ làm việc tiêu chuẩn, giá cước).
  - Một bảng tính tên `Export_LoTrinh` để xuất dữ liệu đã xử lý, với các vùng được đặt tên như `Ngay_Ex`, `TaiXe_Ex`, v.v.
  - Các vùng được đặt tên trong workbook (ví dụ: `data_Export`) để tham chiếu các vùng dữ liệu.

## Cài đặt
1. **Tải hoặc sao chép**: Tải xuống kho lưu trữ này hoặc sao chép các tệp VBA vào máy cục bộ.
2. **Nhập tệp VBA**:
   - Mở workbook Excel.
   - Nhấn `Alt + F11` để mở trình chỉnh sửa VBA.
   - Nhấp chuột phải vào dự án trong Project Explorer, chọn `Import File` và nhập tất cả các tệp `.bas` và `.cls`.
3. **Thiết lập Workbook**:
   - Đảm bảo các bảng tính (`TONG_HOP`, `THONG_TIN_CHUNG`, `Export_LoTrinh`) và bảng (`LoTrinh_Tong`, `ThongTinChung`) được thiết lập như mô tả trong Yêu cầu.
   - Xác định các vùng được đặt tên trong Excel để khớp với các vùng được tham chiếu trong mã (ví dụ: `data_Export`, `Ngay_Ex`).
4. **Bật Macro**: Đảm bảo macro được bật trong Excel để chạy mã VBA.

## Hướng dẫn sử dụng
1. **Chạy thủ tục chính**:
   - Mở workbook Excel.
   - Nhấn `Alt + F8`, chọn `Main` từ danh sách macro và nhấp `Run`.
   - Nhập biển số xe, ngày bắt đầu và ngày kết thúc khi được yêu cầu qua hộp thoại nhập liệu.
2. **Kết quả đầu ra**:
   - Bảng tính `Export_LoTrinh` sẽ được cập nhật với dữ liệu lộ trình, bao gồm các trường tính toán như thời gian làm thêm giờ và quãng đường.
   - Các phép tính tổng hợp (ví dụ: tổng doanh thu, thuế) sẽ được ghi vào các ô được chỉ định trong `Export_LoTrinh`.
   - Một hộp thoại thông báo sẽ xác nhận hoàn tất, cho biết liệu các dòng thừa có bị xóa hay không.
3. **Kiểm tra dữ liệu**:
   - Chạy macro `CheckCellEmpty` để đánh dấu các ô trống trong cột `SoKmDaSuDung` để xem xét.

## Ví dụ quy trình
1. Người dùng chạy macro `Main`.
2. Đầu vào: Biển số xe (`29A-12345`), ngày bắt đầu (`01/05/2025`), ngày kết thúc (`31/05/2025`).
3. Mã thực hiện:
   - Đếm số bản ghi khớp trong `LoTrinh_Tong`.
   - Điều chỉnh số dòng trong `Export_LoTrinh` để khớp với số bản ghi.
   - Trích xuất và ghi dữ liệu lộ trình vào `Export_LoTrinh`.
   - Tính toán thời gian làm thêm giờ và doanh thu, cập nhật các trường tổng hợp.
4. Kết quả: Bảng `Export_LoTrinh` được điền dữ liệu lộ trình và tóm tắt tài chính.

## Lưu ý
- **Hiệu suất**: Mã tắt cập nhật màn hình và tính toán tự động trong quá trình thực thi để cải thiện hiệu suất. Các cài đặt này không được bật lại rõ ràng trong mã hiện tại, vì vậy hãy cân nhắc thêm `Application.ScreenUpdating = True` và `Application.Calculation = xlCalculationAutomatic` vào cuối thủ tục `Main` nếu cần.
- **Xử lý lỗi**: Mã hiện tại giả định các đầu vào và định dạng dữ liệu hợp lệ. Hãy cân nhắc thêm xử lý lỗi cho các ngày không hợp lệ, bảng bị thiếu hoặc đầu vào trống.
- **Tính toàn vẹn dữ liệu**: Đảm bảo bảng `LoTrinh_Tong` và `ThongTinChung` được điền đúng để tránh lỗi runtime.

## Liên hệ
Nếu có thắc mắc hoặc cần hỗ trợ, vui lòng mở một issue trên kho lưu trữ GitHub hoặc liên hệ với người duy trì dự án.