
'AJS

Sub Delete_Rows_Based_On_Value()
'Apply a filter to a Range and delete visible rows

Dim ws As Worksheet

  'Set reference to the sheet in the workbook.
  Set ws = ThisWorkbook.Worksheets("Regular Range")
  ws.Activate 'not required but allows user to view sheet if warning message appears
  
  'Clear any existing filters
  On Error Resume Next
    ws.ShowAllData
  On Error GoTo 0

  '1. Apply Filter
  ws.Range("B3:G1000").AutoFilter Field:=4, Criteria1:=""
  
  '2. Delete Rows
  Application.DisplayAlerts = False
    ws.Range("B4:G1000").SpecialCells(xlCellTypeVisible).Delete
  Application.DisplayAlerts = True
  
  '3. Clear Filter
  On Error Resume Next
    ws.ShowAllData
  On Error GoTo 0

End Sub