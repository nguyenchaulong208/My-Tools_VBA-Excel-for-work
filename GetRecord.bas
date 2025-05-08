Attribute VB_Name = "GetRecord"
Option Explicit
Public dataTbl As Collection
Sub GetRecordFormDatabase()

    Dim dataArr As ThongTinLoTrinh
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim row As ListRow
    Dim tbl As ListObject
    
    Dim cellNgay As Variant
    Dim cellDate As Date
    
    
    
    
    Set wb = ThisWorkbook
    Set ws = wb.Sheets("TONG_HOP")
    Set tbl = ws.ListObjects("LoTrinh_Tong")
    Set dataTbl = New Collection
    
    For Each row In tbl.ListRows
        Set dataArr = New ThongTinLoTrinh
        'Set row = tbl.ListRows(i)
        If dataBsx.bienSoXe_ = row.Range(tbl.ListColumns("BienSoXe").Index).Value Then
            cellNgay = row.Range(tbl.ListColumns("Ngay").Index).Value
            cellDate = DateValue(cellNgay)
            If startDay.ngayBatDau_ <= cellDate And cellDate <= endDay.ngayKetThuc_ Then
                With dataArr
                   .ngay_ = row.Range(tbl.ListColumns("Ngay").Index).Value
                   .diaDiem_ = row.Range(tbl.ListColumns("DiaDiem").Index).Value
                   .thoiGianBd_ = row.Range(tbl.ListColumns("ThoiGianBatDau").Index).Value
                   .thoiGianKt_ = row.Range(tbl.ListColumns("ThoiGianKetThuc").Index).Value
                   .soKmBd_ = row.Range(tbl.ListColumns("SoKmBatDau").Index).Value
                   .soKmKt_ = row.Range(tbl.ListColumns("SoKmKetThuc").Index).Value
                   .quangDuong_ = row.Range(tbl.ListColumns("SoKmDaSuDung").Index).Value
                   .tongTienVe_ = row.Range(tbl.ListColumns("TongTienVetc").Index).Value
                   .soLuongVe_ = row.Range(tbl.ListColumns("SoLuongVe").Index).Value
                   .taiXe_ = row.Range(tbl.ListColumns("TaiXe").Index).Value
                   .bienSoXe_ = row.Range(tbl.ListColumns("BienSoXe").Index).Value
                   .tuyenDuong_ = row.Range(tbl.ListColumns("TuyenDuong").Index).Value
                   .khachHang_ = row.Range(tbl.ListColumns("CongTy").Index).Value
                End With
                dataTbl.Add dataArr
           
             End If
        End If
    
    Next row

End Sub
