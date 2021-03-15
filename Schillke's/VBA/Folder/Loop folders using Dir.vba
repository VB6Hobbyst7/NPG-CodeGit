'TODO: specify folder to loop subfolders for
Dim strFolder As String: strFolder = "C:\temp\"
Dim strItemInFolder As String
strItemInFolder = Dir(strFolder, vbDirectory)
Do While strItemInFolder <> ""
    If ((GetAttr(strFolder & strItemInFolder) And vbDirectory) = vbDirectory) And _
        Not (strItemInFolder = "." Or strItemInFolder = "") Then
        'TODO: replace Debug.Print by the process you want to do on the subfolder
        'Dim strFilePath As String: strFilePath = strFolder & strItemInFolder
        Debug.Print strItemInFolder
    End If
    strItemInFolder = Dir
Loop