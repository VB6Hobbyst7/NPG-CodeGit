'TODO: Specify path
Dim strFolder As String: strFolder = "C:\temp\"
Dim strItemInFolder As String
Dim FolderList() As String 'The array with found folders
Dim intFoundFolders As Integer
strItemInFolder = Dir(strFolder, vbDirectory)
Do While strItemInFolder <> ""
    If ((GetAttr(strFolder & strItemInFolder) And vbDirectory) = vbDirectory) And _
        Not (strItemInFolder = "." Or strItemInFolder = "..") Then
            ReDim Preserve FolderList(intFoundFolders)
            FolderList(intFoundFolders) = strItemInFolder
            intFoundFolders = intFoundFolders + 1
    End If
    strItemInFolder = Dir
Loop