'Attribute VB_Name = "Calculate"
Sub CalRevenue()
    'Khai bao bien chung
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rowTbl As ListRow
    '---------
    'Khai bao bien de tinh toan
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
    Dim soKm As Long
    Dim kmVuot As Long
    Dim soGioOT As Long
    Dim soNgayTangCuong As Long
    '---------
    'Khai bao bien de ghi du lieu va Excel
    Dim slOverTime As Range
    Dim slKm As Range
    Dim writeDoanhThu As Range
    Dim writeTienThue As Range
    Dim writeTongCong As Range
    Dim writeThanhTienCuoc As Range
    Dim writeTangCuong As Range
    Dim writeKmVuot As Range
    Dim writeOverTime As Range
    Dim writeThanhTienVETC As Range
    Dim writeSumVeVETC As Range
    Dim writeSoLuong As Range
    Dim writeDonGiaCuoc As Range
    Dim writeDonGiaKmVuot As Range
    Dim writeDonGiaOverTime As Range
    Dim writeDonGiaTangCuong As Range
    Dim writeSoKmVuot As Range
    Dim writeSoGioOT As Range
    Dim writeSoNgayTangCuong As Range
  
  

    '---------
    'Khai bao bien de lay du lieu tu sheet
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
    set writeSumVeVETC = wss.Names("SumVeVETC_Ex").RefersToRange
    Set writeSoLuong = wss.Names("SumSoLuong_Ex").RefersToRange
    set writeThanhTienCuoc = wss.Names("TT_ThanhTienCuoc_Ex").RefersToRange
    set writeTangCuong = wss.Names("TT_ThanhTienTangCuong_Ex").RefersToRange
    Set writeSoKmVuot = wss.Names("TT_SoKmVuot_Ex").RefersToRange
    set writeKmVuot = wss.Names("TT_ThanhTienKmVuot_Ex").RefersToRange
    set writeOverTime = wss.Names("TT_ThanhTienOverTime_Ex").RefersToRange
    Set writeSoGioOT = wss.Names("TT_OverTime_Ex").RefersToRange
    Set writeThanhTienVETC = wss.Names("TT_ThanhTienVeVETC_Ex").RefersToRange
    Set writeDonGiaCuoc = wss.Names("TT_DonGiaCuoc_Ex").RefersToRange
    Set writeDonGiaKmVuot = wss.Names("TT_DonGiaKmVuot_Ex").RefersToRange
    Set writeDonGiaOverTime = wss.Names("TT_DonGiaOverTime_Ex").RefersToRange
    Set writeSoNgayTangCuong = wss.Names("TT_SLTangCuong_Ex").RefersToRange
    Set writeDonGiaTangCuong = wss.Names("TT_DonGiaChuNhat_Ex").RefersToRange
    
    
    'Duyet kiem tra du lieu va gan cac gia tri
    For Each rowTbl In tbl.ListRows
        If dataBsx.bienSoXe_ = rowTbl.Range(tbl.ListColumns("BienSoXe").Index).Value Then
            'Gan cac gia tri vao bien tuong ung
            donGiaCuoc = rowTbl.Range(tbl.ListColumns("DoanhThuThang").Index).Value
            donGiaChuNhat = rowTbl.Range(tbl.ListColumns("DonGiaNgayChuNhat").Index).Value
            donGiaKmVuot = rowTbl.Range(tbl.ListColumns("DonGiaKmVuot").Index).Value
            donGiaOverTime = rowTbl.Range(tbl.ListColumns("DonGiaQuaGio").Index).Value
            soKm = rowTbl.Range(tbl.ListColumns("KmHopDong").Index).Value
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
    'Tinh toan km vuot
    If sumKm > soKm Then
        kmVuot = sumKm - soKm
    Else
        kmVuot = 0
    End If

    
    
    'Tinh doanh thu
    ttCuocThang = donGiaCuoc * 1
    ttOverTime = (sumOverTime / 60) * donGiaOverTime
    ttVeVETC = sumVeVETC / 1.08
    ttDoanhThu = ttCuocThang + ttOverTime + ttVeVETC + (kmVuot * donGiaKmVuot)+ ttTangCuong
    ttTienThue = ttDoanhThu * 0.08
    ttTongCong = ttDoanhThu + ttTienThue
    soGioOT = sumOverTime / 60
    'Dem so ngay tang cuong
    Dim rowRange as Range
    For Each rowRange in wss.Range("Thu_Ex")
        If rowRange.Value = "Chu Nhat" Then
            soNgayTangCuong = soNgayTangCuong + 1
        End If
    Next rowRange
    'Tinh thanh tien tang cuong
    ttTangCuong = donGiaChuNhat * soNgayTangCuong
   
    'Ghi vao excel
    slOverTime.Value = sumOverTime
    slKm.Value = sumKm
    writeSumVeVETC.value = sumVeVETC
    writeSoLuong.Value = sumSoLuong
    writeThanhTienCuoc.Value = ttCuocThang
    writeTangCuong.Value = ttTangCuong
    writeSoKmVuot.Value = kmVuot
    writeKmVuot.value = kmVuot * donGiaKmVuot
    writeSoGioOT.Value = soGioOT
    writeOverTime.Value = ttOverTime
    writeThanhTienVETC.Value = ttVeVETC
    writeDonGiaCuoc.Value = donGiaCuoc
    writeDonGiaKmVuot.Value = donGiaKmVuot
    writeDonGiaOverTime.Value = donGiaOverTime
    writeDonGiaTangCuong.Value = donGiaChuNhat
    writeDoanhThu.Value = ttDoanhThu
    writeTienThue.Value = ttTienThue
    writeTongCong.Value = ttTongCong

     
     
     
End Sub
