Attribute VB_Name = "GetDataExcel"
Sub GetDataAndExport()
    Dim result As Long
    Dim countData As CountRecord
    Dim insertData As AddData
'Tat nhay man hinh
    Application.ScreenUpdating = False
    ' T?o d?i tu?ng ExportData
    Set countData = New CountRecord
    

    
    result = countData.CountRecordFromTable()

    ' Hi?n th? k?t qu?
    MsgBox "K?t qu?: " & result
  
    
    
End Sub
