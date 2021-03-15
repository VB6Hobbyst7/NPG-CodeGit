Type Record 'TODO: 1. define your fixed lenght record, 2. place the Type block to the top of a module (not a class)
 ID As Integer 
 Name As String * 20 
End Type 
Dim iFile As Integer: iFile = FreeFile
Open <String:Filename> For Input As #iFile 
Dim MyRecord As Record 
Open <String:Filename> For Random As #1 Len = Len(MyRecord) 
Do While Not EOF(iFile)
	Get #iFile,, MyRecord 
	With MyRecord
		<cursor>
	End With
Loop
Close #iFile