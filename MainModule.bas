Attribute VB_Name = "MainModule"
'Function dung de them cac dong du lieu sau khi dem so luong record

Option Explicit
Public sheetName As DataLoTrinh
Sub Main()
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
    
   'Dem so luong dong va them dong
    Application.ScreenUpdating = Fasle
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
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
        
        'Lay du lieu va ghi du lieu vao excel
        Call GetRecordFormDatabase
        Call WriteToExcel
        Call CalRevenue
        MsgBox ("Hoan thanh trich xuat du lieu")
    Else
        'Lay du lieu va ghi du lieu vao excel, Xoa bo dong thua trong Form(Neu co)
        Dim i As Long
        newRowAdd = rowAdd * (-1)
        newRowAdd = Abs(newRowAdd)
         For i = rng.Rows.Count To rng.Rows.Count - newRowAdd + 1 Step -1
            rng.Rows(i).EntireRow.Delete
         Next i
         Call GetRecordFormDatabase
         Call WriteToExcel
         Call CalRevenue
    MsgBox ("Hoan thanh trich xuat du lieu (da xoa dong thua)")
        
    End If
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
  
    

End Sub

