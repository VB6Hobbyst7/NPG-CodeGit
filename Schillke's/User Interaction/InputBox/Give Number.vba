Dim strInput As String
Do
	strInput = InputBox(Prompt:="Give Number")
	If Len(strInput) = 0 Then GoTo ExitGiveNumber
	Dim booIsNumeric As Boolean: booIsNumeric = IsNumeric(strInput)
	If Not booIsNumeric Then
		If vbCancel= MsgBox("You should enter a Number", vbExclamation + vbOkCancel) Then GoTo ExitGiveNumber
	End If
Loop While Not booIsNumeric
ExitGiveNumber: