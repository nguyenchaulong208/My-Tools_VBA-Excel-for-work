Attribute VB_Name = "OverTime"
Function OverTimeFromData(startTime As Date, endTime As Date) As Integer

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
Dim overTimeStart As Double
Dim overTimeEnd As Double

Dim row As ListRow

    Set ws = ThisWorkbook.Worksheets("THONG_TIN_CHUNG")
    Set tbl = ws.ListObjects("ThongTinChung")
    
    For Each row In tbl.ListRows
        If dataBsx.bienSoXe_ = row.Range(tbl.ListColumns("BienSoXe").Index).Value Then
        
            baseStartTime = row.Range(tbl.ListColumns("BatDau").Index).Value
            baseEndTime = row.Range(tbl.ListColumns("KetThuc").Index).Value
        
        End If

    Next row
    ' So sanh voi gio bat dau va ket thuc trong bang thong tin chung
    'overTimeStart = startTime - baseStartTime
    'overTimeEnd = endTime - baseEndTime
    'Tinh gio OverTime
  
    'If startTime < baseStartTime Then
        'Tinh toan va quy doi sang phut
        ' resultStr = (baseStartTime - startTime) * 24 * 60
    'End If
    
    ' If endTime > baseEndTime Then
    '     'Tinh toan va quy doi sang phut
    '     resultEnd = (endTime - baseEndTime) * 24 * 60
    ' End If
    '----------------------------------------
    'Viet lat cach tinh khac

    resultStr = Application.Max(0, (baseStartTime - startTime) * 24 *60)
    resultEnd = Application.Max(0, (endTime - baseEndTime) * 24 *60)
    overTime = resultStr + resultEnd
    'Return gia tri cua OverTime
    OverTimeFromData = overTime
End Function
