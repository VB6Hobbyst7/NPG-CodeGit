Dim iPos As Integer
Dim strExtension As String
iPos = InStrRev(<String:File>, ".")
If iPos <> 0 Then
	strExtension = Mid$(<String:File>, iPos)
End If