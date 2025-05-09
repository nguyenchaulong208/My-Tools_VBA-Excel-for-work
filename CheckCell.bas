Attribute VB_Name = "CheckCell"
Sub CheckCellEmpty()
Dim wb As Workbook
    Dim ws As Worksheet
    Dim cell As Range
    Dim tbl As ListObject
    
    Set wb = ThisWorkbook
    Set ws = wb.Worksheets("TONG_HOP")
    Set tbl = ws.ListObjects("LoTrinh_Tong")
    For Each cell In tbl.ListColumns("SoKmDaSuDung").DataBodyRange
        If Trim(cell.Value) = "" Then
            cell.Interior.Color = RGB(255, 255, 0)
        Else
            cell.Interior.ColorIndex = xlNone
            
        End If
    Next cell
    
End Sub
