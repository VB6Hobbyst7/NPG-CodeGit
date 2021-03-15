Dim strInput As String, strInputLng As String, booNotWholeNumber As Boolean
GiveNumberRetry:
booNotWholeNumber = False
strInput = InputBox("Give Whole Number")
If Len(strInput) = 0 Then GoTo ExitGiveNumber
On Error Resume Next
strInputLng = CLng(strInput)
If Err > 0 Then
    Err = 0
    booNotWholeNumber = True
End If
If strInput <> strInputLng Then booNotWholeNumber = True
If booNotWholeNumber Then
    If vbCancel = MsgBox("You should enter a Whole Number", vbExclamation + vbOKCancel) Then
        GoTo ExitGiveNumber:
    Else
        GoTo GiveNumberRetry
    End If
End If
ExitGiveNumber: