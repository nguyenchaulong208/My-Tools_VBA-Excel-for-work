VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CountRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function CountRecordFromTable() As Long
'Khai bao bien de nhap du lieu
    
    Dim ngayBD As Date
    Dim ngayKt As Date
    Dim bienSoXe As String
'-----
'Khai bao bien xu ly du lieu
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim cell As range
    Dim tbl As ListObject
    Dim CountRecord As Long
    Dim cellDate As Date

'-----
    Set wb = ThisWorkbook
    Set ws = wb.Sheets("TONG_HOP")
    Set tbl = ws.ListObjects("LoTrinh_Tong")
    
'-----
'Nhap thong tin vao Input Box
    bienSoXe = InputBox("Nhap vao bien so xe:", "Input")

    ngayBD = CDate(InputBox("Nhap ngay dau thang(dd/mm/yyyy):", "Input"))
    ngayKt = CDate(InputBox("Nhap ngay cuoi thang (dd/mm/yyyy):", "Input"))

'-----
'Dem so dong thoa dieu kien
CountRecord = 0

For Each cell In tbl.ListColumns("BienSoXe").DataBodyRange
    If bienSoXe = cell.Text Then 'Duyet tim bien so xe
        Set cellNgay = tbl.ListColumns("Ngay").DataBodyRange.Cells(cell.Row - tbl.ListColumns("BienSoXe").DataBodyRange.Row + 1) 'Lay gia tri cot Ngay tuong ung voi hang trong cot
        cellDate = DateValue(cellNgay.Value) 'Chuyen doi ngay trong cot Ngay thanh dinh dang Date
        If IsDate(cellDate) Then 'IsDate: dung de kiem tra neu da chuyen doi thanh cong thi thuc hien buoc tiep theo
                If ngayBD <= cellDate And ngayKt >= cellDate Then
                    CountRecord = CountRecord + 1
                End If
            End If
        End If
    Next cell

CountRecordFromTable = CountRecord
Exit Function

End Function
