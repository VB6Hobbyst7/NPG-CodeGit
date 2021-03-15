<RequiredLevel:3>
Dim FileSpec As String
Dim FileName As String
Dim FileList() As String
Dim FoundFiles As Integer

' Specify path and file spec
FileSpec = "C:\temp\" & "*.*"
FileName = Dir(FileSpec)
' If file is found
If FileName <> "" Then
    FoundFiles = 1
    ReDim FileList(1)
    FileList(FoundFiles) = FileName
	'   Get other filenames
	Do
		FileName = Dir
		If FileName = "" Then Exit Do
		FoundFiles = FoundFiles + 1
		ReDim Preserve FileList(FoundFiles)
		FileList(FoundFiles) = FileName 
	Loop
End If
