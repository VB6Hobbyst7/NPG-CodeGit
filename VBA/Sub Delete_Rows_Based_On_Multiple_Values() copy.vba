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
    lo.DataBodyRange.SpecialCells(xlCellTypeVisible).select.delete
  Application.DisplayAlerts = True

  '3. Clear Filter
  lo.AutoFilter.ShowAllData

End Sub

Sub
Sub LastRowInOneColumn()
   Dim LastRow As Long
   Dim i As Long, j As Long

   'Find the last used row in a Column: column A in this example
   With Worksheets("Sheet2")
      LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
   End With

   MsgBox (LastRow)
   'first row number where you need to paste values in Sheet1'
   With Worksheets("Sheet1")
      j = .Cells(.Rows.Count, "A").End(xlUp).Row + 1
   End With 

   For i = 1 To LastRow
       With Worksheets("Sheet2")
           If .Cells(i, 1).Value = "X" Then
               .Rows(i).Copy Destination:=Worksheets("Sheet1").Range("A" & j)
               j = j + 1
           End If
       End With
   Next i
End Sub