Sub Find_First()
Dim FindString As String
Dim Rng As Range
FindString = "Needs Trained"
If Trim(FindString) <> "" Then
    With Sheets("Sheet1").Range("A:A") 'searches all of column A
        Set Rng = .Find(What:=FindString, _
                        After:=.Cells(.Cells.Count), _
                        LookIn:=xlValues, _
                        LookAt:=xlWhole, _
                        SearchOrder:=xlByRows, _
                        SearchDirection:=xlNext, _
                        MatchCase:=False)
        If Not Rng Is Nothing Then
            Application.Goto Rng, True 'value found
        Else
            MsgBox "Nothing found" 'value not found
        End If
    End With
End If
End Sub
