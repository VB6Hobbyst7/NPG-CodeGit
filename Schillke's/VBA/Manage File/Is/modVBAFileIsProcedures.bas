Attribute VB_Name = "modVBAFileIsProcedures"
Option Explicit
Option Compare Text

Public Function IsFileOpen(FullFileName As String, Optional ErrText As String) As Boolean
'Returns TRUE if the file is open by another process. It returns true
' under the following circumstances (and the ErrText variable will contain the error description).

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim ErrN As Long
Dim ErrD As String
Dim FNum As Integer
If Trim(FullFileName) = vbNullString Then
    IsFileOpen = False
    Exit Function
End If

FNum = FreeFile()
On Error Resume Next
Err.Clear
Open FullFileName For Input Lock Read As #FNum
ErrN = Err.Number
ErrD = Err.Description
On Error Resume Next
Err.Clear
Close #FNum
ErrText = vbNullString
Select Case ErrN
    Case 0
        ' OK
        IsFileOpen = False
    Case 53
        ' File does not exist
        IsFileOpen = False
        ErrText = "File or directory does not exist"
    Case 52
        ' Drive does not exist
        IsFileOpen = False
        ErrText = "Drive does not exist"
    Case 70
        ' File in use "permission denied"
        IsFileOpen = True
End Select

End Function

