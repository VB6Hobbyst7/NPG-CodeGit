Select Case Asc(<String>)
	Case Case 65 To 90, 97 To 122
	<Boolean:IsAlphabetic> = True
	Case Else
	<Boolean:IsAlphabetic> = False
End Select