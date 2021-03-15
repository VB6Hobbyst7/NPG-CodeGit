'AJS
Sub Delete_Rows_Based_On_Value_Table_Message()
'Display Yes/No message prompt before deleting rows

Dim lo As ListObject
Dim lRows As Long
Dim vbAnswer As VbMsgBoxResult

  'Set reference to the sheet and Table.
  Set lo = Sheet6.ListObjects(1)
  lo.Parent.Activate 'Activate sheet that Table is on.
  
  'Clear any existing filters
  lo.AutoFilter.ShowAllData
  
  '1. Apply Filter
  lo.Range.AutoFilter Field:=4, Criteria1:="Product 2"
  
  'Count Rows & display message
  On Error Resume Next
    lRows = WorksheetFunction.Subtotal(103, lo.ListColumns(1).DataBodyRange.SpecialCells(xlCellTypeVisible))
  On Error GoTo 0
  
  vbAnswer = MsgBox(lRows & " Rows will be deleted.  Do you want to continue?", vbYesNo, "Delete Rows Macro")
  
  If vbAnswer = vbYes Then
    
    'Delete Rows
    Application.DisplayAlerts = False
      lo.DataBodyRange.SpecialCells(xlCellTypeVisible).Delete
    Application.DisplayAlerts = True
  
    'Clear Filter
    lo.AutoFilter.ShowAllData
    
  End If

End Sub