Attribute VB_Name = "Preview_Shift"
Sub PreviewShift()
Attribute PreviewShift.VB_ProcData.VB_Invoke_Func = " \n14"
 
' PreviewShift Macro
' This macro will take the selected shift information from the Weekly Summary Page and Link it to the Appropropriate Shifts
' as a preview only.  Supervisors will enter "real" data at the start of their respective shifts

Dim ShiftSelect As String, StatusSelection As String
Dim VehicleID As String, Marker As String, TestType As String
Dim Question As String

'This section identifies the correct shift tabs that will be previewed (1)2/3/1 Dont reset (1)
        


'REFERENCE INFO: Needs updating   (If ColumnNumber = 6 Then ShiftSelect = "Monday 3")



'This section looks for Vehicle Runners on the Shift Summary Page for the selected shift and transfers then to Shift Summary.

RowPosShiftFT = 5
RowPosShiftOthers = 14
ColumnNumber = Range("ColumnNumber")

ShiftSelect = Range("ShiftSelect")


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
            

'Sheets(ShiftSelect).Select  Should be shift 2 of whatever 2/3/1 was processed


End Sub
