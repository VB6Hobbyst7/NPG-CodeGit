Sub LastRowInOneColumn()
   Dim LastRow As Long
   Dim i As Long, j As Long

   'Find the last used row in a Column: column A in this example
   With Worksheets("Training roster")
      LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
   End With

   MsgBox (LastRow)
   'first row number where you need to paste values in Sheet1'
   With Worksheets("Sheet1")
      j = .Cells(.Rows.Count, "A").End(xlUp).Row + 1
   End With 

   For i = 1 To LastRow
       With Worksheets("Training roster")
           If .Cells(i, 1).Value = "X" Then
               .Rows(i).Copy Destination:=Worksheets("Sheet1").Range("A" & j)
               j = j + 1
           End If
       End With
   Next i
End Sub

Dim lastRow As String

lastRow = ActiveSheet.Cells(Rows.Count, "B").End(xlUp).Row + 1
Range("B" & lastRow).Select
Selection.PasteSpecial