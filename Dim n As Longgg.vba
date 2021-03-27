Dim n As Long
n = 5

If n > 0 Then
    'If TRUE then do:
    
Else
    'If FALSE then do:
    
End If
Dim n As Long
n = 5

If n > 0 Then
    'If TRUE then do:
    
ElseIf n = 0 Then
    'Elseif TRUE then do:
    
Else
    'If FALSE then do:
    
End If
Dim n As Long
n = 5

If n < 0 Or n > 10 Then
    'If TRUE then do:
    
End If
Dim n As Long
n = 5

If n > 0 Then
    'If TRUE then do:
    
End If
For Each Cell In Range("$G$41")
    'Do Something
Next Cell
Do Until i < 5
    'Do Something
    i = i + 1
Loop
For i = 1 To 5
    'Do Something
Next i
MyString = "AutomateExcel" 'define string

For Counter = 1 To Len(MyString)
    'do something
Next
For Each ws In Worksheets
    'Do something
Next ws
Sub DeleteMyFile()

On Error Resume Next
    Kill "c:myfile.txt"
'replace this with the directory/file that you want to delete

End Sub
Function DoesFileExist(strFullPath As String) As Boolean
    If Len(Dir(strFullPath)) = 0 Then
        DoesFileExist = False
    Else
        DoesFileExist = True
    End If
End Function
'Get File Name from Path (without extensions)
Function FileNameFromPath_NoExt(strFullPath As String) As String
    Dim nStartLoc As Integer, nEndLoc As Integer, nLength As Integer
    
    nStartLoc = Len(strFullPath) - (Len(strFullPath) - InStrRev(strFullPath, "\") - 1)
    nEndLoc = Len(strFullPath) - (Len(strFullPath) - InStrRev(strFullPath, "."))
    nLength = intEndLoc - intStartLoc
    
    FileNameNoExtensionFromPath = Mid(strFullPath, nStartLoc, nLength)
End Function
Function FileNameFromPath(strFullPath As String) As String
    FileNameFromPath = Right(strFullPath, Len(strFullPath) - InStrRev(strFullPath, "\"))
End Function
Sub List_All_The_Files_Within_Path_Using_Dir()
    Dim nRow As Integer
    Dim strFilePath As String
    Dim strFileName As String
    
    strFilePath = "C:\"
    strFileName = Dir(strFilePath & "*.*", vbReadOnly + vbHidden + vbSystem)
    nRow = 1
    
    'Lists all the files in the current directory
    Do While strFileName <> ""
        Worksheets("Sheet1").Cells(nRow, 2).Value = strFileName
        strFileName = Dir()
        nRow = nRow + 1
    Loop
End Sub
Sub List_All_The_Files_Within_Path_Using_FileSystemObject()
    Dim nRow As Integer
    Dim strFilePath As String
    
    Dim objFile       As Object
    Dim objFSO        As Object
    Dim objFolder     As Object
    Dim objFiles      As Object
    
    strFilePath = "C:\"
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFSO.GetFolder(strFilePath)
    Set objFiles = objFolder.Files
    nRow = 1
    
    'Lists all the files in the current directory
    For Each objFile In objFiles
        Worksheets("Sheet1").Cells(nRow, 1).Value = objFile.name
        nRow = nRow + 1
    Next objFile
End Sub
Public Function SelectFile(Optional strTitle As String = "Select a File", Optional strFilterDescription As String = "All Files", Optional strFilter As String = "*.*", Optional strInitialDirectory As String) As String
    'Usage: Call BrowseFile ("Select a PDF file","PDF Files","*.pdf")
    Dim d As Object
    Set d = Application.FileDialog(3)
    d.AllowMultiSelect = False
    d.Filters.Clear
    d.Title = strTitle
    d.Filters.Add strFilterDescription, strFilter
    d.InitialFileName = strInitialDirectory
    d.Show
    On Error Resume Next
    SelectFile = d.SelectedItems(1)
    Set d = Nothing
End Function

'Usage Example
Private Sub SelectFile_Example()
    Call SelectFile("SelectFile...", "All files", "*.*", "C:\Temp\")
End Sub
Sub Unzip_Files()
    
    'Declare Variables
    Dim oApp As Object
    Dim Fname As Variant
    Dim Output_Folder As Variant
    Dim strDate As String
    Dim i As Long
    
    'Select multiple zip files to unzip
    Fname = Application.GetOpenFilename(filefilter:="Zip Files (*.zip), *.zip", _
      MultiSelect:=True)
    If IsArray(Fname) = False Then
        'Do nothing
    Else
        'Set output folder path for unzip files
        Output_Folder = "C:\Unzip"
        
        'Append backslash to output folder path
        If Right(Output_Folder, 1) <> "\" Then
            Output_Folder = Output_Folder & "\"
        End If
        
        'Extract the files into output folder
        Set oApp = CreateObject("Shell.Application")
        For i = LBound(Fname) To UBound(Fname)
            
            oApp.Namespace(Output_Folder).CopyHere oApp.Namespace(Fname(i)).items
            
        Next i
        
        MsgBox "Files successfully extracted to: " & Output_Folder
    End If
    
End Sub
Function ValidateFileName(ByVal name As String) As Boolean

    ' Check for nothing in filename.
    If name Is Nothing Then
        ValidateFileName = False
    End If

    ' Determines if there are bad characters.
    For Each badChar As Char In System.IO.Path.GetInvalidPathChars
        If InStr(name, badChar) > 0 Then
            ValidateFileName = False
        End If
    Next
    ' If Name passes all above tests Return True.
    ValidateFileName = True
End Function
Public Function Contains(strBaseString As String, strSearchTerm As String) As Boolean
    'Returns TRUE if one string exists within another
    On Error GoTo ErrorMessage
    Debug.Print strBaseString
    Contains = InStr(strBaseString, strSearchTerm)
    Exit Function
ErrorMessage:
    MsgBox "The database has generated an error. Please contact the database administrator, quoting the following error message: '" & Err.Description & "'", vbCritical, "Database Error"
    End
End Function

'Usage Example
Private Sub Contains_Example()
    MsgBox (Contains("TestString", "String"))
End Sub
Public Function CountDelimitedStrings(strMain As String, strDelimiter As String) As Integer
    '* Returns the number of strings in 'strMain' that are delimited by 'strDelimiter'
    '* Ex.: CountDelimitedStrings("abc;def;ghi", ";") returns 3
    Dim i As Integer
    Dim ipos As Integer
    Dim inewpos As Integer
    Dim found As Integer
    Dim Rtn As Integer
    On Error GoTo Err_Section
    
    Rtn = 0
    If IsNull(strMain) Then
    Else
        ipos = 1
        found = True
        Do While found
            inewpos = InStr(ipos, strMain, strDelimiter)
            If inewpos = 0 Then
                found = False
            Else
                If inewpos > ipos Then
                    Rtn = Rtn + 1
                End If
                ipos = inewpos + 1
            End If
        Loop
        If Len(strMain) > (ipos - 1) Then
            Rtn = Rtn + 1
        End If
    End If
    'MsgBox "# of delimited strings = " & rtn
    CountDelimitedStrings = Rtn
    Exit Function
    
Err_Section:
    MsgBox "Error " & Err & " in function CountDelimitedStrings " & Err.Description
    Exit Function
End Function

'Usage Example
Private Sub CountDelimitedStrings_Example()
    MsgBox (CountDelimitedStrings("This;is;a;test", ";"))
End Sub
Public Function EndsWith(str As String, ending As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim endingLen As Integer
    endingLen = Len(ending)
    
    EndsWith = (Right(Trim(UCase(str)), endingLen) = UCase(ending))
    
ProcedureExit:
    Exit Function
    
ErrorHandler:
    MsgBox "Error" & ":  " & Err.Number & vbCrLf & "Description: " _
      & Err.Description, vbExclamation
    Resume ProcedureExit
End Function

'Usage Example
Private Sub EndsWith_Example()
    MsgBox (EndsWith("String", "ing"))
End Sub
Public Function FindString(strCheck As String, strFind As String) As Boolean
    Dim exc As Variant
    Dim Arr
    Dim Flag As Boolean
    
    Arr = Split(strFind, ";")
    Flag = False
    
    For Each exc In Arr
        If Len(exc) > 0 Then
            If InStr(strCheck, exc) > 0 Then
                Flag = True
            End If
        End If
    Next
    FindString = Flag
End Function

'Usage Example
Private Sub FindString_Example()
    MsgBox (FindString("String1String2String3", "String1;String4;String5"))
End Sub
Public Function EndsWith(str As String, ending As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim endingLen As Integer
    endingLen = Len(ending)
    
    EndsWith = (Right(Trim(UCase(str)), endingLen) = UCase(ending))
    
ProcedureExit:
    Exit Function
    
ErrorHandler:
    MsgBox "Error" & ":  " & Err.Number & vbCrLf & "Description: " _
      & Err.Description, vbExclamation
    Resume ProcedureExit
End Function

'Usage Example
Private Sub EndsWith_Example()
    MsgBox (EndsWith("String", "ing"))
End Sub
Public Function GetLastWord(sStr As String) As String
    '* Returns the last word in sStr
    
    Dim i As Integer
    Dim ilen As Integer
    Dim s As String
    Dim sTemp As String
    Dim sLastWord As String
    Dim sHold As String
    Dim iFoundChar As Integer
    
    sTemp = vbNullString
    sLastWord = vbNullString
    iFoundChar = False
    sHold = sStr
    ilen = Len(sStr)
    For i = ilen To 1 Step -1
        s = Right$(sHold, 1)
        If s = " " Then
            If Not iFoundChar Then
                '* skip spaces at end of string.
            Else
                sLastWord = sTemp
                Exit For
            End If
        Else
            iFoundChar = True
            sTemp = s & sTemp
        End If
        If Len(sHold) > 0 Then
            sHold = Left$(sHold, Len(sHold) - 1)
        End If
    Next i
    
    If sLastWord = vbNullString And sTemp <> vbNullString Then
        sLastWord = sTemp
    End If
    GetLastWord = Trim$(sLastWord)
    
End Function

'Usage Example
Private Sub GetLastWord_Example()
    MsgBox (GetLastWord("This is a test"))
End Sub
Public Function GetSubString(strMain As Variant, n As Integer, strDelimiter As String) As String
    
    '* Get the "n"-th substring from "strMain" where strings are delimited by "strDelimiter"
    Dim i As Integer
    Dim substringcount As Integer
    Dim Pos As Integer
    Dim strx As String
    Dim val1 As Integer
    Dim W As String
    
    On Error GoTo Err_GetSubString
    
    GetSubString = vbNullString
    If IsNull(strMain) Then
        Exit Function
    End If
    
    W = vbNullString
    substringcount = 0
    i = 1
    Pos = InStr(i, strMain, strDelimiter)
    Do While Pos <> 0
        strx = Mid$(strMain, i, Pos - i)
        substringcount = substringcount + 1
        If substringcount = n Then
            Exit Do
        End If
        'pddxxx In case the strDelimiter is more than one char
        'i = pos + 1
        i = Pos + Len(strDelimiter)
        Pos = InStr(i, strMain, strDelimiter)
    Loop
    
    If substringcount = n Then
        GetSubString = strx
    Else
        strx = Mid$(strMain, i, Len(strMain) + 1 - i)
        substringcount = substringcount + 1
        If substringcount = n Then
            GetSubString = strx
        Else
            GetSubString = vbNullString
        End If
    End If
    
    On Error GoTo 0
    Exit Function
    
Err_GetSubString:
    MsgBox "GetSubString " & Err & " " & Err.Description
    Resume Next
End Function

'Usage Example
Private Sub GetSubString_Example()
    MsgBox (GetSubString("This,is,a,test", 1, ","))
End Sub
Public Function IsEmailAddress(strEmail As String) As Boolean
    
    On Error GoTo Err:
    
    Dim strURL As String
    
    'Assume is not valid
    IsEmailAddress = False
    
    'Test if valid or not
    With CreateObject("vbscript.regexp")
        .Pattern = "^[\w-\.]+@([\w-]+\.)+[A-Za-z]{2,4}$"
        IsEmailAddress = .test(strEmail)
    End With
    
    If IsEmailAddress And blnCheckDomain Then
        strURL = Right(strEmail, Len(strEmail) - InStr(1, strEmail, "@", vbTextCompare))
        IsEmailAddress = IsURL(strURL)
    End If
    
ExitHere:
    Exit Function
    
Err:
    'If error occured, assume no valid
    IsEmailAddress = False
    GoTo ExitHere
End Function

'Usage Example
Private Sub IsEmailAddress_Example()
    If (IsEmailAddress("mail@gmail.com") = True) Then
        MsgBox ("Email is valid.")
    Else
        MsgBox ("Email is not valid.")
    End If
End Sub
Public Function IsURL(strURL As String) As Boolean
    
    Dim Request As Object
    Dim ff As Integer
    Dim rc As Variant
    
    On Error GoTo EndNow
    
    IsURL = False
    
    Set Request = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    'Add http if not in strURL string
    If Left(strURL, 4) <> "http" Then
        strURL = "http://" & strURL
    End If
    
    With Request
        .Open "GET", strURL, False
        .Send
        rc = .StatusText
    End With
    
    Set Request = Nothing
    
    If rc = "OK" Then IsURL = True
    
EndNow:
    Exit Function
End Function

'Usage Example
Private Sub IsURL_Example()
    If (IsURL("www.google.com") = True) Then
        MsgBox ("URL is valid.")
    Else
        MsgBox ("URL is not valid.")
    End If
End Sub
Public Function RemoveSpaces(strInput As String)
    ' Removes all spaces from a string of text
test:
    If InStr(strInput, " ") = 0 Then
        RemoveSpaces = strInput
    Else
        strInput = Left(strInput, InStr(strInput, " ") - 1) _
          & Right(strInput, Len(strInput) - InStr(strInput, " "))
        GoTo test
    End If
End Function

'Usage Example
Private Sub RemoveSpaces_Example()
    MsgBox (RemoveSpaces("jkl mno pqr"))
End Sub
Public Function RemoveString(strInput As String, StringToRemove As String)
    ' Removes all spaces from a string of text
    Dim s As String
    If InStr(strInput, StringToRemove) = 0 Then
        RemoveString = strInput
    Else
        s = Replace(strInput, StringToRemove, "")
        RemoveString = Replace(s, "  ", " ")
    End If
End Function

'Usage Example
Private Sub RemoveString_Example()
    MsgBox (RemoveString("This is a test string", "test"))
End Sub
Public Function StartsWith(str As String, start As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim startLen As Integer
    startLen = Len(start)
    
    StartsWith = (Left(Trim(UCase(str)), startLen) = UCase(start))
    
ProcedureExit:
    Exit Function
    
ErrorHandler:
    MsgBox "Error" & ":  " & Err.Number & vbCrLf & "Description: " _
      & Err.Description, vbExclamation
    Resume ProcedureExit
End Function

'Usage Example
Private Sub StartsWith_Example()
    MsgBox (StartsWith("String", "Str"))
End Sub
Public Function TrimStringFromEnd(str As String, strToTrim As String) As String
    On Error GoTo ErrorHandler
    
    Dim str1Len As Integer
    Dim str2Len As Integer
    
    str1Len = Len(str)
    str2Len = Len(strToTrim)
    
    If str1Len < str2Len Then  ' We can't trim off a longer string than the original, so just return the original
        TrimStringFromEnd = str
        GoTo ProcedureExit
    End If
    
    If Right(str, str2Len) = strToTrim Then
        ' Trim off str and return the new shorter string
        TrimStringFromEnd = Left(str, str1Len - str2Len)
    Else ' The strToTrim passed in doesn't exist at the end of str, so just return str
        TrimStringFromEnd = str
    End If
    
ProcedureExit:
    Exit Function
    
ErrorHandler:
    MsgBox "Error" & ":  " & Err.Number & vbCrLf & "Description: " _
      & Err.Description, vbExclamation
    Resume ProcedureExit
End Function

'Usage Example
Private Sub TrimStringFromEnd_Example()
    MsgBox (TrimStringFromEnd("TestString", "String"))
End Sub
Function CleanString(ByVal inString As String) As String
    Dim chars As String
    Dim i As Integer
    Dim strip As String
    
    chars = "\/:*""<>|?"
    strip = inString
    
    For i = 1 To Len(chars)
        If InStr(1, strip, Mid(chars, i, 1)) > 0 Then
            strip = Replace(strip, Mid(chars, i, 1), " ")
        End If
    Next i
    
    CleanString = strip
End Function
Public Sub LogToFile(log As String)
    
    ' This subroutine can be used within other subs to log things in a file.
    ' A detailed timestamp will be used, another Function is used for this purpose
    
    On Error GoTo ErrHandler
    
    Const LOG_FILE = "C:\temp\logFile.txt" ' This is the log file fullpath
    
    ' FileSystemObject can be used for file operations
    Dim fs As FileSystemObject
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    ' Getting the folder path of log file
    Dim logFolder As String
    logFolder = fs.GetParentFolderName(LOG_FILE)
    
    ' Check if the folder exists
    If Not fs.FolderExists(logFolder) Then Exit Sub
    
    ' Here we add a timestamp to our log.
    'log = Time_InMS & vbTab & log
    
    ' This is another way to handle files, as well as FileSystemObject
    Dim fNum As Integer
    fNum = FreeFile
    Open LOG_FILE For Append As #fNum
    Print #fNum, log
    Close fNum
    
    Exit Sub
    
ErrHandler:
    
    MsgBox "Error: " & Err.Description, vbCritical, "LogToFile()"
    
End Sub
For Each Cell In Range("$G$41")
    'Do Something
Next Cell
Dim MyInput As String
MyInput = InputBox("Description", "Title", "Default Input Value")
MsgBox "Message Line 1" & vbCrLf & "Message Line 2", , "Message Box Title"
Dim Answer As String
'Display MessageBox - Returns vbYes or vbNo
Answer = MsgBox("Question Description", vbQuestion + vbYesNo, "Title")
Dim Answer As String
'Display MessageBox - Returns vbYes or vbNo
Answer = MsgBox("Question Description", vbQuestion + vbYesNoCancel, "Title")

