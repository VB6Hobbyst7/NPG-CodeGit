Dim strInput As String
Do
    strInput = InputBox("Give Date")
    If Len(strInput) = 0 Then GoTo ExitGiveDate
    If Not IsDate(strInput) Then 
		If vbCancel= MsgBox("You should enter a date", vbExclamation + vbOkCancel) Then GoTo ExitGiveDate
	End If
Loop While Not (IsDate(strInput))
ExitGiveDate: