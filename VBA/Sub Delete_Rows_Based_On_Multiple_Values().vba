Sub Delete_Rows_Based_On_Multiple_Values()
'Apply a filter to a Table and delete visible rows

Dim lo As ListObject

  'Set reference to the sheet and Table.
  Set lo = Sheet5.ListObjects(1)
  lo.Parent.Activate 'Activate sheet that Table is on.
  
  'Clear any existing filters
  lo.AutoFilter.ShowAllData
  
  '1. Apply Filter - Blanks in Product for before 2015 only
  lo.Range.AutoFilter Field:=4, Criteria1:=""
  lo.Range.AutoFilter Field:=1, Criteria1:="<1/1/2015"
  
  '2. Delete Rows
  Application.DisplayAlerts = False
    lo.DataBodyRange.SpecialCells(xlCellTypeVisible).Delete
  Application.DisplayAlerts = True

  '3. Clear Filter
  lo.AutoFilter.ShowAllData

End Sub