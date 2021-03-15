Sub Delete_Rows_Based_On_Value_Table()


Dim lo As ListObject

  'Set reference to the sheet and Table.
  Set lo = Sheet3.ListObjects(1)
  ws.Activate
  
  'Clear any existing filters
  lo.AutoFilter.ShowAllData
  
  '1. Apply Filter
  lo.Range.AutoFilter Field:=4, Criteria1:="Product 2"
  
  '2. Delete Rows
  Application.DisplayAlerts = False
    lo.DataBodyRange.SpecialCells(xlCellTypeVisible).Delete
  Application.DisplayAlerts = True

  '3. Clear Filter
  lo.AutoFilter.ShowAllData

End Sub
