If wsActiveSheet = Sheet24 Then
    'If TRUE then do:
    Call SummaryUpdate
    
Else
    'If FALSE then do:

If Trim(FindString) <> "" Then
    With ActiveSheet.Range("$H$33:$H$55")
        Set Rng = .Find(What:=FindString, _
                        After:=.Cells(.Cells.Count), _
                        LookIn:=xlValues, _
                        LookAt:=xlWhole, _
                        SearchOrder:=xlByRows, _
                        SearchDirection:=xlNext, _
                        MatchCase:=False)
        If Not Rng Is Nothing Then   'If find string is found
            'Call
        Else
           'Call or Do           'If Findstring is not found
        End If
    End With
End If

endmacro:
wsActiveSheet.Protect
Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.EnableEvents = True
End Sub
    
    
    Range("H32:H55").Select
    Selection.AutoFilte
    ActiveSheet.Range("$H$32:$H$55").AutoFilter Field:=1, Criteria1:="<>"
    ActiveSheet.Range("$H$32:$H$55").AutoFilter Field:=1
    
 




End If


'
' Add multi Drivers
'
'
Dim lastRow As String
Dim wsActiveSheet As Worksheet
Dim wsTR As Worksheet
Set wsActiveSheet = ActiveSheet
Set wsTR = Sheets("Training roster")
    
    lastRow = wsTR.Cells(Rows.Count, "A").End(xlUp).Row + 1
        
        Range("K1").Select
            ActiveSheet.Range("$I$4:$I$31").AutoFilter Field:=1, Criteria1:= _
            "Needs Trained"
            Range("S1").Select
            Selection.End(xlDown).Select
            Range(Selection, Selection.End(xlDown)).Select
            Range(Selection, Selection.End(xlToRight)).Select
            Selection.Copy
                Sheets("Training roster").Select
                Range("A" & lastRow).Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
        
        wsActiveSheet.Activate
            Range("K1").Select
            ActiveSheet.Range("$I$4:$I$31").AutoFilter Field:=1
            Application.CutCopyMode = False
                Columns("Q:X").Select
                Selection.EntireColumn.Hidden = True
            wsActiveSheet.Protect
        Range("A1:J1").Select
End Sub

'

'
ct

    Range("H32:H55").Select
    Selection.AutoFilter
    ActiveSheet.Range("$H$32:$H$55").AutoFilter Field:=1, Criteria1:=Array( _
        "In Process", "Not Started", "On Hold"), Operator:=xlFilterValues
    ActiveSheet.ShowAllData
    Selection.AutoFilter
    ActiveWorkbook.Save
    ActiveWindow.SmallScroll Down:=3
    Range("F47").Select
