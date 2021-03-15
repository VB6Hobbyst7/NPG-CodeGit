Attribute VB_Name = "modVBAFolderProcedures"
Option Explicit
Option Compare Text

Private Const MAX_PATH = 260

Public Function MakeDirMulti(DirSpec As String) As Boolean
'Creates multiple nested directories. 
'http://www.cpearson.com/excel/MakeDirMulti.aspx RequiredLevel:3
'This is a replacement function for the VBA MkDir function. MkDir
' will create only the last (right-most) directory of a
' path specification, and all directories to the left of the
' last director must already exist. For example, the following
' will fail
'       MkDir "C:\Folder\Subfolder1\Subfolder2\Subfolder3"
' will fail unless "C:\Folder\Subfolder1\Subfolder2\" already
' exists. MakeDirMulti will create all the folders in
' "C:\Folder\Subfolder1\Subfolder2\Subfolder3" as required.
' If a "\\" string is found, it is converted to "\".
' At present, MakeDirMulti supports local and mapped drives,
' but not UNC paths.
' The function will return True even if no directories were
' created (all directories in DirSpec already existed).
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim Ndx As Long
Dim Arr As Variant
Dim DirString As String
Dim TempSpec As String
Dim DirTestNeeded As Boolean

''''''''''''''''''''''''''''''''
' Ensure DirSpec is valid.
''''''''''''''''''''''''''''''''
If Trim(DirSpec) = vbNullString Then
    MakeDirMulti = False
    Exit Function
End If
If Len(DirSpec) > MAX_PATH Then
    MakeDirMulti = False
    Exit Function
End If
If Not ((Mid(DirSpec, 2, 1) = ":") Or (Mid(DirSpec, 3, 1) = ":")) Then
    MakeDirMulti = False
    Exit Function
End If

'''''''''''''''''''''''''''''''''''''
' Set DirTestNeeded to True. This
' indicates that we need to test to
' see if a folder exists. Once we
' create the first directory, there
' will no longer be a need to call
' Dir to see if a folder exists, since
' the newly created directory will, of
' course, have no existing subfolders.
''''''''''''''''''''''''''''''''''''''
DirTestNeeded = True
TempSpec = DirSpec
'''''''''''''''''''''''''''''''''''''
' If there is a trailing \ character,
' delete it.
'''''''''''''''''''''''''''''''''''''
If Right(TempSpec, 1) = "\" Then
    TempSpec = Left(TempSpec, Len(TempSpec) - 1)
End If

'''''''''''''''''''''''''''''''''
' Split DirSpec into an array,
' delimited by "\".
'''''''''''''''''''''''''''''''''
Arr = Split(expression:=TempSpec, delimiter:="\")
''''''''''''''''''''''''''''''''''''
' Loop through the array, building
' up DirString one folder at a time.
' Each iteration will create
' one directory, moving left to
' right if the folder does not already
' exist.
''''''''''''''''''''''''''''''''''''
For Ndx = LBound(Arr) To UBound(Arr)
    '''''''''''''''''''''''''''''''''
    ' If this is the first iteration
    ' of the loop, just take Arr(Ndx)
    ' without prefixing it with the
    ' existing DirString and path
    ' separator.
    '''''''''''''''''''''''''''''''''
    If Ndx = LBound(Arr) Then
        DirString = Arr(Ndx)
    Else
        DirString = DirString & Application.PathSeparator & Arr(Ndx)
    End If
    On Error GoTo ErrH:
    ''''''''''''''''''''''''''''''''''
    ' Only call the Dir function
    ' if we have yet to create a
    ' new directory. Once we create
    ' a new directory, we no longer
    ' need to call Dir, since the
    ' newly created folder will, of
    ' course, have no subfolders.
    '''''''''''''''''''''''''''''''''
    If DirTestNeeded = True Then
        If Dir(DirString, vbDirectory + vbSystem + vbHidden) = vbNullString Then
            DirTestNeeded = False
            MkDir DirString
        End If
    Else
        MkDir DirString
    End If
    On Error GoTo 0
Next Ndx

MakeDirMulti = True
Exit Function

ErrH:
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' If an error occured, typically because an invalid
' character was encountered in a directory name, return
' False.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
MakeDirMulti = False
End Function

