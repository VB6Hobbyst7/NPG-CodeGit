Attribute VB_Name = "Update_Summary"
Sub SummaryUpdate()
Attribute SummaryUpdate.VB_ProcData.VB_Invoke_Func = " \n14"
'
' SummaryUpdate Macro
'This macro will take the vehicle that ran from the Individual Shift Sheet to the Overall Summary Page

Dim ShiftSelect As String, StatusSelection As String
Dim VehicleID As String, Marker As String, TestType As String
Dim SheetName As String, DriverIndicator As String
Dim StatusAtShiftEnd As String, Rank As String

Dim ColumnNumber As Integer
Dim CellText As String, VarSplit As Variant

' Selects the Shift that is to be updated on the Summary Sheet by looking at the tab name.
SheetName = ActiveSheet.Name

If SheetName = "Monday 3" Then ColumnNumber = 6
If SheetName = "Monday 1" Then ColumnNumber = 7
If SheetName = "Monday 2" Then ColumnNumber = 8
If SheetName = "Tuesday 3" Then ColumnNumber = 9
If SheetName = "Tuesday 1" Then ColumnNumber = 10
If SheetName = "Tuesday 2" Then ColumnNumber = 11
If SheetName = "Wednesday 3" Then ColumnNumber = 12
If SheetName = "Wednesday 1" Then ColumnNumber = 13
If SheetName = "Wednesday 2" Then ColumnNumber = 14
If SheetName = "Thursday 3" Then ColumnNumber = 15
If SheetName = "Thursday 1" Then ColumnNumber = 16
If SheetName = "Thursday 2" Then ColumnNumber = 17
If SheetName = "Friday 3" Then ColumnNumber = 18
If SheetName = "Friday 1" Then ColumnNumber = 19
If SheetName = "Friday 2" Then ColumnNumber = 20
If SheetName = "Saturday 3" Then ColumnNumber = 21
If SheetName = "Saturday 1" Then ColumnNumber = 22
If SheetName = "Saturday 2" Then ColumnNumber = 23
If SheetName = "Sunday 3" Then ColumnNumber = 24
If SheetName = "Sunday 1" Then ColumnNumber = 25
If SheetName = "Sunday 2" Then ColumnNumber = 26
If SheetName = "New Monday 3" Then ColumnNumber = 28
If SheetName = "New Monday 1" Then ColumnNumber = 29


'This clears the previous selections from the Summary Sheet Shift Column that is being updated.
'This is performed because not everything on the list might have ran.

Sheets("Priorities & Summary").Select
  For J = 7 To 45
    Cells(J, ColumnNumber) = ""
  Next J


'This section looks for actual vehicle runners in the Field Test Group section

For I = 5 To 12
  Sheets(SheetName).Select
  DriverIndicator = Cells(I, 5).Value
    If DriverIndicator <> "" Then
        If DriverIndicator <> "Did Not Run" Then
            If DriverIndicator <> "Did Not Arrive" Then
                If DriverIndicator <> "No resources" Then
                    VehicleID = Cells(I, 2).Value
                    Sheets("Priorities & Summary").Select
                        For J = 7 To 14
                            If Cells(J, 3).Value = VehicleID Then
                                Cells(J, ColumnNumber) = "x"
                                J = 45
                            End If
                        Next J
                End If
            End If
       End If
    End If
Next I
         

'This section looks for actual vehicle runners in the NON Field Test Group section

For I = 14 To 31
  Sheets(SheetName).Select
  DriverIndicator = Cells(I, 5).Value
    If DriverIndicator <> "" Then
        If DriverIndicator <> "Did Not Run" Then
            If DriverIndicator <> "Did Not Arrive" Then
                If DriverIndicator <> "No resources" Then
                    VehicleID = Cells(I, 2).Value
                    Sheets("Priorities & Summary").Select
                        For J = 15 To 45
                            If Cells(J, 3).Value = VehicleID Then
                                Cells(J, ColumnNumber) = "x"
                                J = 45
                            End If
                        Next J
                End If
            End If
        End If
    End If
Next I


'This Section considers whether the vehicle is Down at end of shift.  If Down, it adds a D on future planned shifts of the vehicle.
'It first looks at Field Test Vehicles

For I = 5 To 12
  Sheets(SheetName).Select
  VehicleID = Cells(I, 2).Value
    If VehicleID <> "" Then
       StatusAtShiftEnd = Cells(I, 8).Value
          'Consider a check here to see if Status at end of shift is blank
       If StatusAtShiftEnd = "Down" Or StatusAtShiftEnd = "down" Then
          Sheets("Priorities & Summary").Select
            For J = 7 To 14
              If Cells(J, 3).Value = VehicleID Then
                 'Loop to check for future shifts and add D as prefix
                 For S = 1 To 29 - ColumnNumber
                    Rank = Cells(J, ColumnNumber + S).Value
                    If InStr(1, Rank, "D") Or Rank = "" Then
                       Else
                       Cells(J, ColumnNumber + S).Value = "D-" & Rank
                    End If
                 Next S
              End If
          Next J
       End If
    End If
Next I

'Next, it looks at all other Vehicles

For I = 14 To 31
  Sheets(SheetName).Select
  VehicleID = Cells(I, 2).Value
    If VehicleID <> "" Then
       StatusAtShiftEnd = Cells(I, 8).Value
          'Consider a check here to see if Status at end of shift is blank
       If StatusAtShiftEnd = "Down" Or StatusAtShiftEnd = "down" Then
          Sheets("Priorities & Summary").Select
            For J = 15 To 45
              If Cells(J, 3).Value = VehicleID Then
                 'Loop to check for future shifts and add D as prefix
                 For S = 1 To 29 - ColumnNumber
                    Rank = Cells(J, ColumnNumber + S).Value
                    If InStr(1, Rank, "D") Or Rank = "" Then
                       Else
                       Cells(J, ColumnNumber + S).Value = "D-" & Rank
                    End If
                 Next S
              End If
          Next J
       End If
    End If
Next I


'This Section considers whether the vehicle is Up at end of shift.
'If Up, it looks to see if vehicle was previously down.  If Yes, it will remove the D- from the future shifts.
'It first looks at Field Test Vehicles

For I = 5 To 12
  Sheets(SheetName).Select
  VehicleID = Cells(I, 2).Value
    If VehicleID <> "" Then
       StatusAtShiftEnd = Cells(I, 8).Value
          'Consider a check here to see if Status at end of shift is blank
       If StatusAtShiftEnd = "Up" Or StatusAtShiftEnd = "up" Then
          Sheets("Priorities & Summary").Select
            For J = 7 To 14
              If Cells(J, 3).Value = VehicleID Then
                 'Loop to check for future shifts that have D as prefix and remove the D-
                  For S = 1 To 29 - ColumnNumber
                      CellText = Cells(J, ColumnNumber + S).Value
                      If InStr(1, CellText, "D-") Then
                         VarSplit = Split(CellText, "-", 2)
                         strPart1 = VarSplit(0)
                         strPart2 = VarSplit(1)
                         Cells(J, ColumnNumber + S).Value = strPart2
                     End If
                  Next S
              End If
          Next J
       End If
    End If
Next I


'Next it looks at Other Test Vehicles for Up at end of shift

For I = 14 To 31
  Sheets(SheetName).Select
  VehicleID = Cells(I, 2).Value
    If VehicleID <> "" Then
       StatusAtShiftEnd = Cells(I, 8).Value
          'Consider a check here to see if Status at end of shift is blank
       If StatusAtShiftEnd = "Up" Or StatusAtShiftEnd = "up" Then
          Sheets("Priorities & Summary").Select
            For J = 15 To 45
              If Cells(J, 3).Value = VehicleID Then
                 'Loop to check for future shifts that have D as prefix and remove the D-
                  For S = 1 To 29 - ColumnNumber
                      CellText = Cells(J, ColumnNumber + S).Value
                      If InStr(1, CellText, "D-") Then
                         VarSplit = Split(CellText, "-", 2)
                         strPart1 = VarSplit(0)
                         strPart2 = VarSplit(1)
                         Cells(J, ColumnNumber + S).Value = strPart2
                     End If
                  Next S
              End If
          Next J
       End If
    End If
Next I

Sheets("Priorities & Summary").Select
    
End Sub
