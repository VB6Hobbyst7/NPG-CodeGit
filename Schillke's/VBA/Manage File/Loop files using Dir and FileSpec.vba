Dim strFileName As String
'TODO: Specify path and file spec
Dim strFolder As String: strFolder = "C:\temp\"
Dim strFileSpec As String: strFileSpec = strFolder & "*.*"
strFileName = Dir(strFileSpec)
Do While Len(strFileName) > 0
    'TODO: replace Debug.Print by the process you want to do on the file
    'Dim strFilePath As String: strFilePath = strFolder & strFileName
    Debug.Print strFileName
    strFileName = Dir
Loop