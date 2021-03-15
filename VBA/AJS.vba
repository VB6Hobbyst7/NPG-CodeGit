Sub Delete_Rows_Based_On_Value_Table()

Sub AddDrivers()

'
' Add Drivers
'

'
Dim FindString As String
Dim Rng As Range
FindString = "Needs Trained"
If Trim(FindString) <> "" Then
    With ActiveSheet.Range("$I$4:$I$31") 'searches all of column A
        Set Rng = .Find(What:=FindString, _
                        After:=.Cells(.Cells.Count), _
                        LookIn:=xlValues, _
                        LookAt:=xlWhole, _
                        SearchOrder:=xlByRows, _
                        SearchDirection:=xlNext, _
                        MatchCase:=False)
        If Not Rng Is Nothing Then
            Call AddDrivers
        Else
            MsgBox "No Drivers Need Trained" 'value not found
        End If
    End With
End If
End Sub
 
        Range("K1").Select
        ActiveSheet.Range("$I$4:$I$31").AutoFilter Field:=1, Criteria1:= _
            "Needs Trained"
        Range("S1").Select
        Selection.End(xlDown).Select
        Range(Selection, Selection.End(xlDown)).Select
        Range(Selection, Selection.End(xlToRight)).Select
        Selection.Copy
            Sheets("Training roster").Select
            Range("A2").Select
            Selection.End(xlDown).Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Sheets(Active).Select
        Range("K1").Select
        ActiveSheet.Range("$I$4:$I$31").AutoFilter Field:=1
        Application.CutCopyMode = False
        Range("A1:J1").Select
