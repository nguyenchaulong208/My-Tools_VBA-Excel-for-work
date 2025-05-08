Attribute VB_Name = "WriteData"
Sub WriteToExcel()

    Dim ws As Worksheet
    
    Dim ngayThang As Range
    Dim taiXe As Range
    Dim diaDiem As Range
    Dim startTime As Range
    Dim endTime As Range
    Dim overTime As Range
    Dim km As Range
    Dim veVETC As Range
    Dim soLuong As Range
    Dim nextRow As Long
    Dim item As ThongTinLoTrinh
        
    
    
    
    
    Set ws = ThisWorkbook.Worksheets(sheetName.tenSheet_)
    Set ngayThang = ws.Names("Ngay_Ex").RefersToRange
    Set taiXe = ws.Names("TaiXe_Ex").RefersToRange
    Set diaDiem = ws.Names("DiaDiem_Ex").RefersToRange
    Set startTime = ws.Names("StartTime_Ex").RefersToRange
    Set endTime = ws.Names("EndTime_Ex").RefersToRange
    Set overTime = ws.Names("OverTime_Ex").RefersToRange
    Set km = ws.Names("KM_Ex").RefersToRange
    Set veVETC = ws.Names("VeVETC_Ex").RefersToRange
    Set soLuong = ws.Names("SoLuong_Ex").RefersToRange
    
    nextRow = ngayThang.row 'dong dau tien cua vung
    
    
    
    For Each item In dataTbl
        item.ngay_ = Format(CDate(item.ngay_), "dd/mm/yyyy")
        ws.Cells(nextRow, ngayThang.Column).Value = item.ngay_
        ws.Cells(nextRow, taiXe.Column).Value = item.taiXe_
        ws.Cells(nextRow, diaDiem.Column).Value = item.diaDiem_
        ws.Cells(nextRow, startTime.Column).Value = item.thoiGianBd_
        ws.Cells(nextRow, endTime.Column).Value = item.thoiGianKt_
        ws.Cells(nextRow, overTime.Column).Value = OverTimeFromData(item.thoiGianBd_, item.thoiGianKt_)
        ws.Cells(nextRow, km.Column).Value = item.quangDuong_
        ws.Cells(nextRow, veVETC.Column).Value = item.tongTienVe_
        ws.Cells(nextRow, soLuong.Column).Value = item.soLuongVe_
        
        nextRow = nextRow + 1
    Next item
    
        
    
    

End Sub
