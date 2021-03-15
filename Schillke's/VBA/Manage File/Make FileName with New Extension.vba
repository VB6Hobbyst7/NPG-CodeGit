Dim iPos As Integer
Dim strExtension As String
Dim strExtensionNew As String: strExtensionNew = "" 'TODO: specify new extension 
Dim strFileNameNew As String
iPos = InStrRev(<String:File>, ".")
If iPos <> 0 Then
	strFileNameNew  = Left(<String:File>, iPos) & strExtensionNew 
End If