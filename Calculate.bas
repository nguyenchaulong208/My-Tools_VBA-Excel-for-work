Attribute VB_Name = "Calculate"
Sub CalRevenue()

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rowTbl As ListRow
    Dim donGiaCuoc As Long
    Dim donGiaChuNhat As Long
    Dim donGiaKmVuot As Long
    Dim donGiaOverTime As Long
    Dim ttCuocThang As Long
    Dim ttTangCuong As Long
    Dim ttKmVuot As Long
    Dim ttGioVuot As Long
    Dim ttDoanhThu As Long
    Dim ttTienThue As Long
    Dim ttTongCong As Long
    Dim ttVeVETC As Long
    Dim sumOverTime As Long
    Dim sumKm As Long
    Dim sumVeVETC As Long
    Dim sumSoLuong As Long
    Dim slOverTime As Range
    Dim slKm As Range
    Dim writeDoanhThu As Range
    Dim writeTienThue As Range
    Dim writeTongCong As Range
    Dim wss As Worksheet
   
    Set ws = ThisWorkbook.Worksheets("THONG_TIN_CHUNG")
    Set wss = ThisWorkbook.Worksheets("Export_LoTrinh")
    Set tbl = ws.ListObjects("ThongTinChung")
    'Set vi tri ghi du lieu
    Set slOverTime = wss.Names("SumOverTime_Ex").RefersToRange
    Set slKm = wss.Names("SumKM_Ex").RefersToRange
    Set writeDoanhThu = wss.Names("TT_TongThanhTien_Ex").RefersToRange
    Set writeTienThue = wss.Names("TT_TienThue_Ex").RefersToRange
    Set writeTongCong = wss.Names("TT_TongCong_Ex").RefersToRange
    
    
    
    For Each rowTbl In tbl.ListRows
        If dataBsx.bienSoXe_ = rowTbl.Range(tbl.ListColumns("BienSoXe").Index).Value Then
            'Gan cac gia tri vao bien tuong ung
            donGiaCuoc = rowTbl.Range(tbl.ListColumns("DoanhThuThang").Index).Value
            donGiaChuNhat = rowTbl.Range(tbl.ListColumns("DonGiaNgayChuNhat").Index).Value
            donGiaKmVuot = rowTbl.Range(tbl.ListColumns("DonGiaKmVuot").Index).Value
            donGiaOverTime = rowTbl.Range(tbl.ListColumns("DonGiaQuaGio").Index).Value
        'Else
            'donGiaCuoc = 0
            'donGiaChuNhat = 0
            'donGiaKmVuot = 0
            'donGiaOverTime = 0
        End If

    Next rowTbl
    
    'Tinh toan tong tren bang lo trinh
    sumOverTime = Application.WorksheetFunction.Sum(wss.Range("OverTime_Ex"))
    sumKm = Application.WorksheetFunction.Sum(wss.Range("Km_Ex"))
    sumVeVETC = Application.WorksheetFunction.Sum(wss.Range("VeVETC_Ex"))
    sumSoLuong = Application.WorksheetFunction.Sum(wss.Range("SoLuong_Ex"))
    
    'Tinh doanh thu
    ttCuocThang = donGiaCuoc * 1
    ttOverTime = (sumOverTime / 60) * donGiaOverTime
    ttVeVETC = sumVeVETC / 1.08
    ttDoanhThu = ttCuocThang + ttOverTime + ttVeVETC
    ttTienThue = ttDoanhThu * 0.08
    ttTongCong = ttDoanhThu + ttTienThue
    
    'Ghivao excel
     slOverTime.Value = sumOverTime
     slKm.Value = sumKm
     writeDoanhThu.Value = ttDoanhThu
     writeTienThue.Value = ttTienThue
     writeTongCong.Value = ttTongCong
     
     
     
End Sub
