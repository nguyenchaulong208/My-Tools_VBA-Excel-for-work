Attribute VB_Name = "AddRow"
'Function dung de them cac dong du lieu sau khi dem so luong record

Option Explicit
Public sheetName As DataLoTrinh
Sub AddRowNameRange()
    Dim rng As Range
    Dim wb As Workbook
    Dim nameObject As String
    Dim df As Name
    Dim countRNR As Long
    Dim callCountRecord As Long
    Dim rowAdd As Long
    Dim idObject As String
    Dim cell As Range
    Dim newRowAdd As Long
    Dim row As ListRow
    Dim ws As Worksheet
    Dim inputWorkSheet As String
    
    Dim nextRow As Long
    

    Set wb = ThisWorkbook
    Set ws = wb.Worksheets("Export_LoTrinh")

    Set sheetName = New DataLoTrinh
    sheetName.tenSheet_ = "Export_LoTrinh"
    
    callCountRecord = CountRecordFromTable()
   
 
        
    
    Set df = ThisWorkbook.Names("data_Export")
    Set rng = df.RefersToRange
    countRNR = rng.Rows.Count
    
    
    rowAdd = callCountRecord - countRNR
    If rowAdd > 0 Then
        ws.Rows(rng.row + 1).Resize(rowAdd).Insert Shift:=xlDown
        Call GetRecordFormDatabase
        Call WriteToExcel
        MsgBox ("Hoan thanh trich xuat du lieu")
    Else
        Dim i As Long
        newRowAdd = rowAdd * (-1)
        newRowAdd = Abs(newRowAdd)
         For i = rng.Rows.Count To rng.Rows.Count - newRowAdd + 1 Step -1
            rng.Rows(i).EntireRow.Delete
         Next i
         Call GetRecordFormDatabase
         Call WriteToExcel
    MsgBox ("Hoan thanh trich xuat du lieu (da xoa dong thua)")
        
    End If

End Sub

