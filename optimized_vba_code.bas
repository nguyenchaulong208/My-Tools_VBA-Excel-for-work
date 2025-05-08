'========================================================================
' Optimized VBA Code for Route Data Processing
' 
' Description: This VBA application extracts vehicle route data based on
' license plate and date range, processes it, and outputs results.
'========================================================================

'========================================================================
' CLASS MODULE: DataLoTrinh
'========================================================================
Option Explicit

Private bienSoXe As String
Private ngayBatDau As Date
Private ngayKetThuc As Date
Private tenSheet As String

Public Property Let bienSoXe_(ByRef newBienSoXe As String)
    bienSoXe = newBienSoXe
End Property
Public Property Get bienSoXe_() As String
    bienSoXe_ = bienSoXe
End Property

Public Property Let ngayBatDau_(newNgayBatDau As Date)
    ngayBatDau = newNgayBatDau
End Property
Public Property Get ngayBatDau_() As Date
    ngayBatDau_ = ngayBatDau
End Property

Public Property Let ngayKetThuc_(newNgayKetThuc As Date)
    ngayKetThuc = newNgayKetThuc
End Property
Public Property Get ngayKetThuc_() As Date
    ngayKetThuc_ = ngayKetThuc
End Property

Public Property Let tenSheet_(newTenSheet As String)
    tenSheet = newTenSheet
End Property
Public Property Get tenSheet_() As String
    tenSheet_ = tenSheet
End Property

'========================================================================
' CLASS MODULE: ThongTinLoTrinh
'========================================================================
Option Explicit

Private ngay As String
Private diaDiem As String
Private thoiGianBd As String
Private thoiGianKt As String
Private lamThemGio As Integer
Private soKmBd As Long
Private soKmKt As Long
Private quangDuong As Long
Private tongTienVe As Long
Private soLuongVe As Integer
Private bienSoXe As String
Private taiXe As String
Private khachHang As String
Private tuyenDuong As String

'Properties for ngay
Public Property Let ngay_(ByRef newNgay As String)
    ngay = newNgay
End Property
Public Property Get ngay_() As String
    ngay_ = ngay
End Property

'Properties for diaDiem
Public Property Let diaDiem_(ByRef newDiaDiem As String)
    diaDiem = newDiaDiem
End Property
Public Property Get diaDiem_() As String
    diaDiem_ = diaDiem
End Property

'Properties for thoiGianBd
Public Property Let thoiGianBd_(ByRef newTGBD As String)
    thoiGianBd = newTGBD
End Property
Public Property Get thoiGianBd_() As String
    thoiGianBd_ = thoiGianBd
End Property

'Properties for thoiGianKt
Public Property Let thoiGianKt_(ByRef newTGKT As String)
    thoiGianKt = newTGKT
End Property
Public Property Get thoiGianKt_() As String
    thoiGianKt_ = thoiGianKt
End Property

'Properties for soKmBd
Public Property Let soKmBd_(ByVal newKmBd As Long)
    soKmBd = newKmBd
End Property
Public Property Get soKmBd_() As Long
    soKmBd_ = soKmBd
End Property

'Properties for soKmKt
Public Property Let soKmKt_(ByVal newKmKt As Long)
    soKmKt = newKmKt
End Property
Public Property Get soKmKt_() As Long
    soKmKt_ = soKmKt
End Property

'Properties for quangDuong
Public Property Let quangDuong_(ByVal newQD As Long)
    quangDuong = newQD
End Property
Public Property Get quangDuong_() As Long
    quangDuong_ = quangDuong
End Property

'Properties for tongTienVe
Public Property Let tongTienVe_(ByVal newTien As Long)
    tongTienVe = newTien
End Property
Public Property Get tongTienVe_() As Long
    tongTienVe_ = tongTienVe
End Property

'Properties for soLuongVe
Public Property Let soLuongVe_(ByVal newSLV As Integer)
    soLuongVe = newSLV
End Property
Public Property Get soLuongVe_() As Integer
    soLuongVe_ = soLuongVe
End Property

'Properties for bienSoXe
Public Property Let bienSoXe_(ByRef newBSX As String)
    bienSoXe = newBSX
End Property
Public Property Get bienSoXe_() As String
    bienSoXe_ = bienSoXe
End Property

'Properties for taiXe
Public Property Let taiXe_(ByRef newTaiXe As String)
    taiXe = newTaiXe
End Property
Public Property Get taiXe_() As String
    taiXe_ = taiXe
End Property

'Properties for khachHang
Public Property Let khachHang_(ByRef newKH As String)
    khachHang = newKH
End Property
Public Property Get khachHang_() As String
    khachHang_ = khachHang
End Property

'Properties for tuyenDuong
Public Property Let tuyenDuong_(ByRef newTD As String)
    tuyenDuong = newTD
End Property
Public Property Get tuyenDuong_() As String
    tuyenDuong_ = tuyenDuong
End Property

'Properties for lamThemGio
Public Property Let lamThemGio_(ByRef newLamThemGio As Integer)
    lamThemGio = newLamThemGio
End Property
Public Property Get lamThemGio_() As Integer
    lamThemGio_ = lamThemGio
End Property

'========================================================================
' MODULE: Main
'========================================================================
Option Explicit

' Global variables
Public dataBsx As DataLoTrinh
Public startDay As DataLoTrinh
Public endDay As DataLoTrinh
Public dataTbl As Collection
Public sheetName As DataLoTrinh

' Main entry point
Sub ProcessVehicleRouteData()
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Count records and prepare data
    Dim recordCount As Long
    recordCount = CountRecordFromTable()
    
    If recordCount = 0 Then
        MsgBox "Không tìm thấy dữ liệu nào với các tiêu chí đã nhập", vbInformation, "Thông báo"
        GoTo CleanUp
    End If
    
    ' Add or remove rows as needed
    AddRowNameRange
    
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    MsgBox "Đã xảy ra lỗi: " & Err.Description, vbCritical, "Lỗi"
    
CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub

'========================================================================
' MODULE: CountRecord
'========================================================================
Public Function CountRecordFromTable() As Long
    On Error GoTo ErrorHandler
    
    ' Declare variables
    Dim ngayBD As Date, ngayKt As Date
    Dim bienSoXe As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim cell As Range, cellNgay As Range
    Dim tbl As ListObject
    Dim CountRecord As Long
    Dim cellDate As Date
    
    ' Initialize objects
    Set wb = ThisWorkbook
    Set ws = wb.Sheets("TONG_HOP")
    Set tbl = ws.ListObjects("LoTrinh_Tong")
    Set dataBsx = New DataLoTrinh
    Set startDay = New DataLoTrinh
    Set endDay = New DataLoTrinh
    
    ' Get user input
    bienSoXe = InputBox("Nhập vào biển số xe:", "Input")
    If bienSoXe = "" Then
        MsgBox "Không có biển số xe nào được nhập!", vbExclamation
        CountRecordFromTable = 0
        Exit Function
    End If
    
    dataBsx.bienSoXe_ = bienSoXe
    
    On Error Resume Next
    ngayBD = CDate(InputBox("Nhập ngày đầu tháng (dd/mm/yyyy):", "Input"))
    If Err.Number <> 0 Then
        MsgBox "Định dạng ngày không hợp lệ!", vbExclamation
        CountRecordFromTable = 0
        Exit Function
    End If
    
    ngayKt = CDate(InputBox("Nhập ngày cuối tháng (dd/mm/yyyy):", "Input"))
    If Err.Number <> 0 Then
        MsgBox "Định dạng ngày không hợp lệ!", vbExclamation
        CountRecordFromTable = 0
        Exit Function
    End If
    On Error GoTo ErrorHandler
    
    ' Validate date range
    If ngayKt < ngayBD Then
        MsgBox "Ngày kết thúc phải sau ngày bắt đầu!", vbExclamation
        CountRecordFromTable = 0
        Exit Function
    End If
    
    startDay.ngayBatDau_ = ngayBD
    endDay.ngayKetThuc_ = ngayKt
    
    ' Count matching records
    CountRecord = 0
    
    Dim dataRange As Range
    Set dataRange = tbl.ListColumns("BienSoXe").DataBodyRange
    
    ' Optimize search by using Find method first
    Set cell = dataRange.Find(What:=bienSoXe, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not cell Is Nothing Then
        Dim firstAddress As String
        firstAddress = cell.Address
        
        Do
            Dim rowIndex As Long
            rowIndex = cell.Row - dataRange.Row + 1
            
            Set cellNgay = tbl.ListColumns("Ngay").DataBodyRange.Cells(rowIndex)
            
            On Error Resume Next
            cellDate = DateValue(cellNgay.Value)
            If Err.Number = 0 Then
                If ngayBD <= cellDate And ngayKt >= cellDate Then
                    CountRecord = CountRecord + 1
                End If
            End If
            On Error GoTo ErrorHandler
            
            Set cell = dataRange.FindNext(cell)
        Loop Until cell.Address = firstAddress
    End If
    
    CountRecordFromTable = CountRecord
    Exit Function
    
ErrorHandler:
    MsgBox "Lỗi khi đếm bản ghi: " & Err.Description, vbCritical, "Lỗi"
    CountRecordFromTable = 0
End Function

'========================================================================
' MODULE: AddRow
'========================================================================
Sub AddRowNameRange()
    On Error GoTo ErrorHandler
    
    Dim rng As Range
    Dim wb As Workbook
    Dim df As Name
    Dim countRNR As Long
    Dim callCountRecord As Long
    Dim rowAdd As Long
    Dim newRowAdd As Long
    Dim ws As Worksheet
    
    Set wb = ThisWorkbook
    Set ws = wb.Worksheets("Export_LoTrinh")
    
    Set sheetName = New DataLoTrinh
    sheetName.tenSheet_ = "Export_LoTrinh"
    
    callCountRecord = CountRecordFromTable()
    
    ' If no records found, exit
    If callCountRecord = 0 Then
        Exit Sub
    End If
    
    Set df = ThisWorkbook.Names("data_Export")
    Set rng = df.RefersToRange
    countRNR = rng.Rows.Count
    
    rowAdd = callCountRecord - countRNR
    
    ' Add or remove rows as needed
    If rowAdd > 0 Then
        ' Add rows
        ws.Rows(rng.Row + 1).Resize(rowAdd).Insert Shift:=xlDown
        GetRecordFormDatabase
        WriteToExcel
        MsgBox "Hoàn thành trích xuất dữ liệu", vbInformation, "Thông báo"
    ElseIf rowAdd < 0 Then
        ' Remove extra rows
        Dim i As Long
        newRowAdd = Abs(rowAdd)
        
        For i = rng.Rows.Count To rng.Rows.Count - newRowAdd + 1 Step -1
            rng.Rows(i).EntireRow.Delete
        Next i
        
        GetRecordFormDatabase
        WriteToExcel
        MsgBox "Hoàn thành trích xuất dữ liệu (đã xóa dòng thừa)", vbInformation, "Thông báo"
    Else
        ' Same number of rows, just update data
        GetRecordFormDatabase
        WriteToExcel
        MsgBox "Hoàn thành trích xuất dữ liệu", vbInformation, "Thông báo"
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Lỗi khi điều chỉnh số dòng: " & Err.Description, vbCritical, "Lỗi"
End Sub

'========================================================================
' MODULE: GetRecord
'========================================================================
Sub GetRecordFormDatabase()
    On Error GoTo ErrorHandler
    
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
    
    ' Create filter array for optimization
    Dim licenseFilter As Collection
    Set licenseFilter = New Collection
    
    ' First filter by license plate
    For Each row In tbl.ListRows
        If dataBsx.bienSoXe_ = row.Range(tbl.ListColumns("BienSoXe").Index).Value Then
            ' Now check date range
            On Error Resume Next
            cellNgay = row.Range(tbl.ListColumns("Ngay").Index).Value
            cellDate = DateValue(cellNgay)
            
            If Err.Number = 0 Then
                If startDay.ngayBatDau_ <= cellDate And cellDate <= endDay.ngayKetThuc_ Then
                    Set dataArr = New ThongTinLoTrinh
                    
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
                    
                    On Error GoTo ErrorHandler
                    dataTbl.Add dataArr
                End If
            End If
            On Error GoTo ErrorHandler
        End If
    Next row
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Lỗi khi lấy dữ liệu: " & Err.Description, vbCritical, "Lỗi"
End Sub

'========================================================================
' MODULE: OverTime
'========================================================================
Function OverTimeFromData(startTime As Date, endTime As Date) As Integer
    On Error GoTo ErrorHandler
    
    Dim result As Date
    Dim ws As Worksheet
    Dim rng As Range
    Dim rngName As Name
    Dim tbl As ListObject
    Dim baseStartTime As Date
    Dim baseEndTime As Date
    Dim resultStr As Double
    Dim resultEnd As Double
    Dim overTime As Integer
    Dim row As ListRow
    
    ' Default values in case something goes wrong
    resultStr = 0
    resultEnd = 0
    
    ' Get base times from ThongTinChung table
    Set ws = ThisWorkbook.Worksheets("THONG_TIN_CHUNG")
    Set tbl = ws.ListObjects("ThongTinChung")
    
    ' Find the matching vehicle record
    For Each row In tbl.ListRows
        If dataBsx.bienSoXe_ = row.Range(tbl.ListColumns("BienSoXe").Index).Value Then
            baseStartTime = row.Range(tbl.ListColumns("BatDau").Index).Value
            baseEndTime = row.Range(tbl.ListColumns("KetThuc").Index).Value
            Exit For
        End If
    Next row
    
    ' Calculate overtime before standard start time
    If IsDate(startTime) And IsDate(baseStartTime) Then
        If startTime < baseStartTime Then
            resultStr = (baseStartTime - startTime) * 24 * 60
        End If
    End If
    
    ' Calculate overtime after standard end time
    If IsDate(endTime) And IsDate(baseEndTime) Then
        If endTime > baseEndTime Then
            resultEnd = (endTime - baseEndTime) * 24 * 60
        End If
    End If
    
    ' Total overtime in minutes
    overTime = resultStr + resultEnd
    OverTimeFromData = overTime
    
    Exit Function
    
ErrorHandler:
    MsgBox "Lỗi khi tính giờ làm thêm: " & Err.Description, vbCritical, "Lỗi"
    OverTimeFromData = 0
End Function

'========================================================================
' MODULE: WriteData
'========================================================================
Sub WriteToExcel()
    On Error GoTo ErrorHandler
    
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
    
    ' Initialize objects and ranges
    Set ws = ThisWorkbook.Worksheets(sheetName.tenSheet_)
    
    ' Use more error-resistant approach for getting ranges
    On Error Resume Next
    Set ngayThang = ws.Names("Ngay_Ex").RefersToRange
    Set taiXe = ws.Names("TaiXe_Ex").RefersToRange
    Set diaDiem = ws.Names("DiaDiem_Ex").RefersToRange
    Set startTime = ws.Names("StartTime_Ex").RefersToRange
    Set endTime = ws.Names("EndTime_Ex").RefersToRange
    Set overTime = ws.Names("OverTime_Ex").RefersToRange
    Set km = ws.Names("KM_Ex").RefersToRange
    Set veVETC = ws.Names("VeVETC_Ex").RefersToRange
    Set soLuong = ws.Names("SoLuong_Ex").RefersToRange
    
    If ngayThang Is Nothing Then
        MsgBox "Không tìm thấy vùng dữ liệu mục tiêu trong sheet " & sheetName.tenSheet_, vbCritical, "Lỗi"
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    
    nextRow = ngayThang.Row
    
    ' Batch operations for better performance
    Application.ScreenUpdating = False
    
    ' Write data to Excel
    For Each item In dataTbl
        ' Format date for consistency
        item.ngay_ = Format(CDate(item.ngay_), "dd/mm/yyyy")
        
        ' Write cell values
        ws.Cells(nextRow, ngayThang.Column).Value = item.ngay_
        ws.Cells(nextRow, taiXe.Column).Value = item.taiXe_
        ws.Cells(nextRow, diaDiem.Column).Value = item.diaDiem_
        ws.Cells(nextRow, startTime.Column).Value = item.thoiGianBd_
        ws.Cells(nextRow, endTime.Column).Value = item.thoiGianKt_
        
        ' Calculate overtime
        On Error Resume Next
        ws.Cells(nextRow, overTime.Column).Value = OverTimeFromData(item.thoiGianBd_, item.thoiGianKt_)
        On Error GoTo ErrorHandler
        
        ws.Cells(nextRow, km.Column).Value = item.quangDuong_
        ws.Cells(nextRow, veVETC.Column).Value = item.tongTienVe_
        ws.Cells(nextRow, soLuong.Column).Value = item.soLuongVe_
        
        nextRow = nextRow + 1
    Next item
    
    ' Restore screen updating
    Application.ScreenUpdating = True
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Lỗi khi ghi dữ liệu vào Excel: " & Err.Description, vbCritical, "Lỗi"
End Sub