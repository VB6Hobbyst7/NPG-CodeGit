Attribute VB_Name = "Update_Shift"
Sub UpdateShift()
Attribute UpdateShift.VB_ProcData.VB_Invoke_Func = " \n14"
'
' UpdateShift Macro
' This macro will take the selected shift information from the Weekly Summary Page and Link it to the Appropropriate Shift Summary
'
Dim ShiftSelect As String, StatusSelection As String
Dim VehicleID As String, Marker As String, TestType As String
Dim Question As String

' Selects the Shift that is wanting to be summerized by looking at the Column Number location.
Sheets("Priorities & Summary").Select
ColumnNumber = ActiveCell.Column

'This section identifies the correct shift tab that will be updated based on column number position
        
If ColumnNumber = 6 Then ShiftSelect = "Monday 3"
If ColumnNumber = 7 Then ShiftSelect = "Monday 1"
If ColumnNumber = 8 Then ShiftSelect = "Monday 2"
If ColumnNumber = 9 Then ShiftSelect = "Tuesday 3"
If ColumnNumber = 10 Then ShiftSelect = "Tuesday 1"
If ColumnNumber = 11 Then ShiftSelect = "Tuesday 2"
If ColumnNumber = 12 Then ShiftSelect = "Wednesday 3"
If ColumnNumber = 13 Then ShiftSelect = "Wednesday 1"
If ColumnNumber = 14 Then ShiftSelect = "Wednesday 2"
If ColumnNumber = 15 Then ShiftSelect = "Thursday 3"
If ColumnNumber = 16 Then ShiftSelect = "Thursday 1"
If ColumnNumber = 17 Then ShiftSelect = "Thursday 2"
If ColumnNumber = 18 Then ShiftSelect = "Friday 3"
If ColumnNumber = 19 Then ShiftSelect = "Friday 1"
If ColumnNumber = 20 Then ShiftSelect = "Friday 2"
If ColumnNumber = 21 Then ShiftSelect = "Saturday 3"
If ColumnNumber = 22 Then ShiftSelect = "Saturday 1"
If ColumnNumber = 23 Then ShiftSelect = "Saturday 2"
If ColumnNumber = 24 Then ShiftSelect = "Sunday 3"
If ColumnNumber = 25 Then ShiftSelect = "Sunday 1"
If ColumnNumber = 26 Then ShiftSelect = "Sunday 2"
If ColumnNumber = 28 Then ShiftSelect = "New Monday 3"
If ColumnNumber = 29 Then ShiftSelect = "New Monday 1"

'Insertion of Message Box to insure correct SHIFT/COLUMN is highlighted.  If yes, then run program.

Range(Cells(7, ColumnNumber), Cells(45, ColumnNumber)).Select
Question = "Are you sure you want to update " & ShiftSelect & "?"
MSG1 = MsgBox(Question, vbOKCancel, "Confirm Date and Shift?")

If MSG1 = vbOK Then
  
'This section clears the previous selections from the Shift that will be updated.

Sheets(ShiftSelect).Select
For I = 5 To 12
  Cells(I, 1).Value = ""
  Cells(I, 2).Value = ""
  Cells(I, 3).Value = ""
  Cells(I, 4).Value = ""
Next I

For I = 14 To 31
  Cells(I, 1).Value = ""
  Cells(I, 2).Value = ""
Next I

'This section looks for Vehicle Runners on the Shift Summary Page for the selected shift and transfers then to Shift Summary.

RowPosShiftFT = 5
RowPosShiftOthers = 14

For I = 7 To 45
    Sheets("Priorities & Summary").Select
    StatusSelection = Cells(I, ColumnNumber).Value
        If StatusSelection <> "" And StatusSelection <> "H" And StatusSelection <> "C" And StatusSelection <> "*" Then
           VehicleID = Cells(I, 3).Value
           Marker = Cells(I, ColumnNumber).Value
           TestType = Cells(I, 1).Value
           Sheets(ShiftSelect).Select
              If TestType = "FT" Then
                 Cells(RowPosShiftFT, 2).Value = VehicleID
                 Cells(RowPosShiftFT, 1).Value = Marker
                 RowPosShiftFT = RowPosShiftFT + 1
                 Else
                   Cells(RowPosShiftOthers, 2).Value = VehicleID
                   Cells(RowPosShiftOthers, 1).Value = Marker
                   RowPosShiftOthers = RowPosShiftOthers + 1
                  
               End If
         End If
Next I
            

Sheets(ShiftSelect).Select

'This section looks for actual vehicle runners in the Field Test Group and highlights cells YELLOW
ActiveSheet.Unprotect
For I = 5 To 12
  VehicleID = Cells(I, 2).Value
    If VehicleID <> "" Then
       Cells(I, 5).Select
         With Selection.Interior
             .Pattern = xlSolid
             .PatternColorIndex = xlAutomatic
             .Color = 65535
             .TintAndShade = 0
             .PatternTintAndShade = 0
         End With
        Cells(I, 8).Select
         With Selection.Interior
             .Pattern = xlSolid
             .PatternColorIndex = xlAutomatic
             .Color = 65535
             .TintAndShade = 0
             .PatternTintAndShade = 0
         End With
    End If
Next I

'This section looks for actual vehicle runners in the Non-Field Test Group and highlights cells YELLOW
For I = 14 To 31
  VehicleID = Cells(I, 2).Value
    If VehicleID <> "" Then
       Cells(I, 5).Select
         With Selection.Interior
             .Pattern = xlSolid
             .PatternColorIndex = xlAutomatic
             .Color = 65535
             .TintAndShade = 0
             .PatternTintAndShade = 0
         End With
        Cells(I, 8).Select
         With Selection.Interior
             .Pattern = xlSolid
             .PatternColorIndex = xlAutomatic
             .Color = 65535
             .TintAndShade = 0
             .PatternTintAndShade = 0
         End With
    End If
Next I
Cells(5, 5).Select

'This section looks for Work Requests and highlights cells YELLOW
For I = 33 To 55
  VehicleID = Cells(I, 2).Value
    If VehicleID <> "" Then
       Cells(I, 5).Select
         With Selection.Interior
             .Pattern = xlSolid
             .PatternColorIndex = xlAutomatic
             .Color = 65535
             .TintAndShade = 0
             .PatternTintAndShade = 0
         End With
       Cells(I, 8).Select
         With Selection.Interior
             .Pattern = xlSolid
             .PatternColorIndex = xlAutomatic
             .Color = 65535
             .TintAndShade = 0
             .PatternTintAndShade = 0
         End With

    End If
Next I
Cells(5, 5).Select
ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
ActiveWorkbook.RefreshALL

'If Shift/Column Slected was incorrect
Else
  MsgBox "Hit okay, move cursor to appropriate column and re-run Shift Update?"
    
End If

'
End Sub
