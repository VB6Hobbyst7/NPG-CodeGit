Attribute VB_Name = "basErrorHandle"

Option Explicit

'Set address to where support emails should be sent to here... (support@mycompany.com)
Private Const AddressTo As String = "write your support email address here"

'reading registry
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Private Const VER_NT_WORKSTATION = 1&
Private Const STANDARD_RIGHTS_READ = &H20000
Private Const KEY_QUERY_VALUE = &H1&
Private Const KEY_ENUMERATE_SUB_KEYS = &H8&
Private Const KEY_NOTIFY = &H10&
Private Const Synchronize = &H100000
Private Const MAXLEN = 256
Private Const ERROR_SUCCESS = &H0&
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or _
                        KEY_QUERY_VALUE Or _
                        KEY_ENUMERATE_SUB_KEYS Or _
                        KEY_NOTIFY) And _
                        (Not Synchronize))
Private Const REG_SZ = 1
Private Const REG_BINARY = 3
Private Const REG_DWORD = 4
'mail
Private Const RET_OK As Long = 0
Private Const RET_FAIL As Long = vbObjectError - 503
#If VBA7 Then
Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" _
      Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation _
      As String, ByVal lpFile As String, ByVal lpParameters As String, _
      ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
#Else
Private Declare Function ShellExecute Lib "shell32.dll" _
      Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation _
      As String, ByVal lpFile As String, ByVal lpParameters As String, _
      ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
#End If
'os version
Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2
Private Type OSVERSIONINFO
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
   szCSDVersion As String * 128   ' Maintenance string for PSS usage.
End Type
  
'OSVERSIONINFOEX supported on NT4 or later only, so a test is required before using
Private Type OSVERSIONINFOEX
  OSVSize            As Long
  dwVerMajor         As Long
  dwVerMinor         As Long
  dwBuildNumber      As Long
  PlatformID         As Long
  szCSDVersion       As String * 128
  wServicePackMajor  As Integer
  wServicePackMinor  As Integer
  wSuiteMask         As Integer
  wProductType       As Byte
  wReserved          As Byte
End Type

#If VBA7 Then
Declare PtrSafe Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
  (lpVersionInformation As Any) As Long
#Else
Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
  (lpVersionInformation As Any) As Long
#End If

#If VBA7 Then
Private Declare PtrSafe Function apiRegOpenKeyEx Lib "advapi32.dll" _
        Alias "RegOpenKeyExA" _
        (ByVal hKey As Long, _
        ByVal lpSubKey As String, _
        ByVal ulOptions As Long, _
        ByVal samDesired As Long, _
        ByRef phkResult As Long) _
        As Long
Private Declare PtrSafe Function apiRegQueryInfoKey Lib "advapi32.dll" _
        Alias "RegQueryInfoKeyA" _
        (ByVal hKey As Long, _
        ByVal lpClass As String, _
        ByRef lpcbClass As Long, _
        ByVal lpReserved As Long, _
        ByRef lpcSubKeys As Long, _
        ByRef lpcbMaxSubKeyLen As Long, _
        ByRef lpcbMaxClassLen As Long, _
        ByRef lpcValues As Long, _
        ByRef lpcbMaxValueNameLen As Long, _
        ByRef lpcbMaxValueLen As Long, _
        ByRef lpcbSecurityDescriptor As Long, _
        ByRef lpftLastWriteTime As FILETIME) _
        As Long
Private Declare PtrSafe Function apiRegQueryValueEx Lib "advapi32.dll" _
        Alias "RegQueryValueExA" _
        (ByVal hKey As Long, _
        ByVal lpValueName As String, _
        ByVal lpReserved As Long, _
        ByRef lpType As Long, _
        lpData As Any, _
        ByRef lpcbData As Long) _
        As Long
Private Declare PtrSafe Function apiRegCloseKey Lib "advapi32.dll" _
        Alias "RegCloseKey" _
        (ByVal hKey As Long) _
        As Long
'getting tempdir
Private Declare PtrSafe Function apiGetTempDir Lib "kernel32" _
        Alias "GetTempPathA" (ByVal nBufferLength As Long, _
        ByVal lpBuffer As String) As Long
#Else
Private Declare Function apiRegOpenKeyEx Lib "advapi32.dll" _
        Alias "RegOpenKeyExA" _
        (ByVal hKey As Long, _
        ByVal lpSubKey As String, _
        ByVal ulOptions As Long, _
        ByVal samDesired As Long, _
        ByRef phkResult As Long) _
        As Long
Private Declare Function apiRegQueryInfoKey Lib "advapi32.dll" _
        Alias "RegQueryInfoKeyA" _
        (ByVal hKey As Long, _
        ByVal lpClass As String, _
        ByRef lpcbClass As Long, _
        ByVal lpReserved As Long, _
        ByRef lpcSubKeys As Long, _
        ByRef lpcbMaxSubKeyLen As Long, _
        ByRef lpcbMaxClassLen As Long, _
        ByRef lpcValues As Long, _
        ByRef lpcbMaxValueNameLen As Long, _
        ByRef lpcbMaxValueLen As Long, _
        ByRef lpcbSecurityDescriptor As Long, _
        ByRef lpftLastWriteTime As FILETIME) _
        As Long
Private Declare Function apiRegQueryValueEx Lib "advapi32.dll" _
        Alias "RegQueryValueExA" _
        (ByVal hKey As Long, _
        ByVal lpValueName As String, _
        ByVal lpReserved As Long, _
        ByRef lpType As Long, _
        lpData As Any, _
        ByRef lpcbData As Long) _
        As Long
Private Declare Function apiRegCloseKey Lib "advapi32.dll" _
        Alias "RegCloseKey" _
        (ByVal hKey As Long) _
        As Long
'getting tempdir
Private Declare Function apiGetTempDir Lib "kernel32" _
        Alias "GetTempPathA" (ByVal nBufferLength As Long, _
        ByVal lpBuffer As String) As Long
#End If


Private Const MAX_PATH As Integer = 255

'if logging is required, add the mErrorLogging module manually
Sub ErrorHandle(Err As ErrObject, _
                    ErrLine As Long, _
                    Optional strProcedure As String, _
                    Optional strComment As String, _
                    Optional bShowMessage As Boolean = True, _
                    Optional bReportError As Boolean = True, _
                    Optional bLogError As Boolean = True)
'Global error handling procedure
'http://codevba.com/help/error_handling.htm#Errorreportingusingmail
Dim strDescription As String
Dim ErrNo As Long
Dim strSource As String
Dim strExtendedErrInfo As String
Dim strProductVersion As String
    
    With Err
        ErrNo = .Number
        strDescription = .Description
        strSource = .Source
    End With
       
    If bShowMessage Then ErrorMessage ErrNo, _
                        ErrLine, _
                        strDescription, _
                        strComment, _
                        strSource, _
                        strProcedure, _
                        bReportError, _
                        strExtendedErrInfo

    If bLogError Then
'implement error logging here
    End If
HandleExit:
End Sub

'Sub ErrorMessageTest()
'    ErrorMessage 1, 0, "Description", "Comment", "Source", "Procedure"
'End Sub

Private Sub ErrorMessage(ErrNo As Long, _
                        ErrLine As Long, _
                        strDescription As String, _
                        Optional strComment As String, _
                        Optional strSource As String, _
                        Optional strProcedure As String, _
                        Optional bReportError As Boolean = True, _
                        Optional strExtendedErrInfo As String = "")
Const cstrError As String = "Error"
Dim strOfficeApplication As String
Dim strDocument As String
Dim strErrorTitle As String
Dim strMessage As String
    Dim strSubject As String
'ExtendedErrInfo only sent by mail

On Error Resume Next
    If Len(strComment) > 0 Then strComment = vbCrLf & vbCrLf & strComment
    If Len(strProcedure) > 0 Or Len(strSource) > 0 Then strProcedure = vbCrLf & " in " & _
                                                        strSource & "." & strProcedure
    
    Dim app As Object: Set app = Application

    strOfficeApplication = app.Name & " (" & app.Version & ")"

    Select Case app.Name
    Case "Microsoft Excel"
        strDocument = app.ThisWorkbook.Name
    Case "Microsoft Access"
        strDocument = app.CodeProject.Name
    End Select
    
    'don't add Error to title/msg if it is a non-reportable > less frightening
    If bReportError = True Then
        strErrorTitle = cstrError
        strMessage = cstrError
    End If
       
    strMessage = strMessage & " " & ErrNo & ": " & strDescription & " " & strProcedure & _
                                        " line " & ErrLine & " " & strComment
    If bReportError = False Then
        MsgBox strMessage, vbCritical, strErrorTitle
    Else
        Dim iPos As Long
        iPos = InStr(strMessage, "@")
        If iPos > 0 Then strMessage = Left(strMessage, iPos - 1)
        Dim lngRet As Long
        Dim strMsg As String
        strMsg = strMessage & vbCrLf & "Note: To speed up our support it would be very helpful if you can sent us the database in which the error occurs, if possible (zipped)" & vbCrLf & _
                    vbCrLf & "DESCRIBE WHAT YOU TRY TO DO: " & vbCrLf & vbCrLf & vbCrLf & _
                    "DESCRIBE WHAT HAPPENS: " & vbCrLf & vbCrLf & vbCrLf & _
                    strDocument & vbCrLf & _
                    WindowsVersion & vbCrLf & _
                    strOfficeApplication & vbCrLf & _
                    strExtendedErrInfo

        If CheckForOLEMessaging() = True Then
           If (vbYes = MsgBox(strMessage & vbCrLf & vbCrLf & "Do you want to report the problem?", vbYesNo + vbCritical + vbDefaultButton2, strErrorTitle)) Then
                Send AddressTo, strErrorTitle, strMsg
           End If
        Else
            Dim strReportFile As String
            Dim intFn As Long
            strReportFile = DirTemporary() & "~" & format(Now, "YYYYMMDDHHNNSS") & ".txt"
            intFn = FreeFile
            Open strReportFile For Output Access Write As #intFn
            Print #intFn, strMsg
            Close #intFn
            Shell "notepad.exe """ & strReportFile & """"
            Kill strReportFile
        End If
    End If

End Sub

Private Function CheckForOLEMessaging() As Boolean
On Error GoTo errHandle
   Dim bOK As Boolean
   CheckForOLEMessaging = (RegistryValueGet(&H80000002, "SOFTWARE\Microsoft\Windows Messaging Subsystem\", "OleMessaging", bOK) = "1")
errHandle:
End Function

Private Function DirTemporary() As String
'Returns Temp Folder Name
Dim strTemp As String
Dim lngRtn As Long
On Error GoTo HandleErr
    strTemp = String$(MAX_PATH, 0)
    lngRtn = apiGetTempDir(MAX_PATH, strTemp)
    If lngRtn <> 0 Then
        DirTemporary = Left$(strTemp, lngRtn)
    Else
        DirTemporary = ""
    End If
HandleExit:
    Exit Function
HandleErr:
    Resume HandleExit
End Function

Private Function RegistryValueGet(ByVal lngKeyToGet As Long, _
                            ByVal strKeyName As String, _
                            ByVal strValueName As String, _
                            bOK As Boolean) _
                            As String

Dim lnghKey As Long
Dim strClassName As String
Dim lngClassLen As Long
Dim lngReserved As Long
Dim lngSubKeys As Long
Dim lngMaxSubKeyLen As Long
Dim lngMaxClassLen As Long
Dim lngValues As Long
Dim lngMaxValueNameLen As Long
Dim lngMaxValueLen As Long
Dim lngSecurity As Long
Dim ftLastWrite As FILETIME
Dim lngType As Long
Dim lngData As Long
Dim lngTmp As Long
Dim strRet As String
Dim varRet As Variant
Dim lngRet As Long
    On Error GoTo RegistryValueGet_Err
    'Open the key first
    lngTmp = apiRegOpenKeyEx(lngKeyToGet, strKeyName, 0&, KEY_READ, lnghKey)

    'Are we ok?
    If Not (lngTmp = ERROR_SUCCESS) Then Err.Raise lngTmp + vbObjectError

    lngReserved = 0&
    strClassName = String$(MAXLEN, 0):  lngClassLen = MAXLEN
    'Get boundary values
    lngTmp = apiRegQueryInfoKey(lnghKey, strClassName, _
        lngClassLen, lngReserved, lngSubKeys, lngMaxSubKeyLen, _
        lngMaxClassLen, lngValues, lngMaxValueNameLen, _
        lngMaxValueLen, lngSecurity, ftLastWrite)

    'How we doin?
    If Not (lngTmp = ERROR_SUCCESS) Then Err.Raise lngTmp + vbObjectError

    'Now grab the value for the key
    strRet = String$(MAXLEN - 1, 0)
    lngTmp = apiRegQueryValueEx(lnghKey, strValueName, _
                lngReserved, lngType, ByVal strRet, lngData)
    Select Case lngType
        Case REG_SZ
            lngTmp = apiRegQueryValueEx(lnghKey, strValueName, _
                lngReserved, lngType, ByVal strRet, lngData)
            varRet = Left(strRet, lngData - 1)
        Case REG_DWORD
            lngTmp = apiRegQueryValueEx(lnghKey, strValueName, _
                lngReserved, lngType, lngRet, lngData)
            varRet = lngRet
        Case REG_BINARY
            lngTmp = apiRegQueryValueEx(lnghKey, strValueName, _
                lngReserved, lngType, ByVal strRet, lngData)
            varRet = Left(strRet, lngData)
    End Select

    'All quiet on the western front?
    If Not (lngTmp = ERROR_SUCCESS) Then
        Err.Raise lngTmp + vbObjectError
    Else
        bOK = True
    End If
    
RegistryValueGet_Exit:
    RegistryValueGet = varRet
    lngTmp = apiRegCloseKey(lnghKey)
    Exit Function
RegistryValueGet_Err:
    varRet = vbNullString
    Resume RegistryValueGet_Exit
End Function


'                   Win95     Win98     WinNT 3.51     WinNT 4.0   Win2000 Win XP   VISTA   Win 7    Win 8
'                  ---------------------------------------------------------------------------------------
'dwPlatFormID         1         1            2            2        2        2       2        2        2
'
'dwMajorVersion       4         4            3            4        5        5       6        6        6
'
'dwMinorVersion       0        10           51            0        0        1       0        1        2
'WIN32s=WIN32s or WinNT <3.51

Private Function WindowsVersion(Optional lngPlatformId As Long, _
                       Optional lngMajorVersion As Long, _
                       Optional lngMinorVersion As Long, _
                       Optional lngBuildNumber As Long _
                        ) As String
'retrieve basic systeminfo
Dim v As OSVERSIONINFO, retval As Long
Dim strWindowsVersion As String, BuildVersion As String
Const cDiv As String = "/"
    On Error GoTo HandleErr
    v.dwOSVersionInfoSize = Len(v)
    retval = GetVersionEx(v)
    
    lngPlatformId = v.dwPlatformId
    lngMajorVersion = v.dwMajorVersion
    lngMinorVersion = v.dwMinorVersion
    lngBuildNumber = v.dwBuildNumber
    
    strWindowsVersion = v.dwMajorVersion & "." & v.dwMinorVersion
    BuildVersion = v.dwBuildNumber And &HFFFF& 'mu?
    
    Select Case v.dwPlatformId
    Case VER_PLATFORM_WIN32_WINDOWS
        Select Case v.dwMinorVersion
        Case 0
            WindowsVersion = "Windows 95"
        Case 10
            WindowsVersion = "Windows 98"
        End Select
    Case VER_PLATFORM_WIN32_NT
        Select Case v.dwMajorVersion
        Case 3
            WindowsVersion = "WinNT 3.51"
        Case 4
            WindowsVersion = "WinNT 4"
        Case 5
            If v.dwMinorVersion = 0 Then
                WindowsVersion = "Windows 2000"
            ElseIf v.dwMinorVersion = 1 Then
                WindowsVersion = "Windows XP"
            ElseIf v.dwMinorVersion = 2 Then
                WindowsVersion = "Windows Server 2003"
            End If
        Case 6
            Dim osvex As OSVERSIONINFOEX
            osvex.OSVSize = Len(osvex)
            retval = GetVersionEx(osvex)
            
            If v.dwMinorVersion = 0 Then
                If osvex.wProductType = VER_NT_WORKSTATION Then
                    WindowsVersion = "Windows Vista"
                Else
                    WindowsVersion = "Windows Server 2008"
                End If
            ElseIf v.dwMinorVersion = 1 Then
                If osvex.wProductType = VER_NT_WORKSTATION Then
                    WindowsVersion = "Windows 7"
                Else
                    WindowsVersion = "Windows Server 2008 R2"
                End If
                WindowsVersion = "Windows 7"
            ElseIf v.dwMinorVersion = 2 Then
                If osvex.wProductType = VER_NT_WORKSTATION Then
                    WindowsVersion = "Windows 8"
                Else
                    WindowsVersion = "Windows Server 2012"
                End If
            ElseIf v.dwMinorVersion = 3 Then
                'For applications that have been manifested for Windows 8.1.
                'Applications not manifested for 8.1 will return the Windows 8 OS version value (6.2).
                If osvex.wProductType = VER_NT_WORKSTATION Then
                    WindowsVersion = "Windows 8.1"
                Else
                    WindowsVersion = "Windows Server 2012 R2"
                End If
            End If
        End Select
    Case VER_PLATFORM_WIN32s
        WindowsVersion = "NT < 3.51"
    End Select

HandleExit:
    Exit Function
HandleErr:
    WindowsVersion = "SysInfo not available"
    Resume HandleExit
End Function

Private Function Send(ByVal vstrAddrTo As String, _
                ByVal vstrSubject As String, _
                ByVal vstrBodyText As String, _
                Optional ByVal vstrAddrCC As String = "", _
                Optional ByVal vstrAddrBCC As String = "", _
                Optional ByVal vfTruncateUntilLastVbCrlf As Boolean = False) As Long

On Error GoTo HandleErr
Dim strMsg As String
Dim lngRet As Long
Dim lngIdx As Long
    strMsg = "mailto:" & vstrAddrTo & "?"
    If Len(vstrAddrCC) > 0 Then
        strMsg = strMsg & "CC=" & vstrAddrCC & "&"
    End If
    If Len(vstrAddrBCC) > 0 Then
        strMsg = strMsg & "BCC=" & vstrAddrBCC & "&"
    End If
    
    vstrSubject = Replace(vstrSubject, vbCrLf, "%0d%0a")
    vstrSubject = Replace(vstrSubject, " ", "%20")
    strMsg = strMsg & "Subject=" & vstrSubject & "&"
    
    vstrBodyText = Replace(vstrBodyText, vbCrLf, "%0d%0a")
    vstrBodyText = Replace(vstrBodyText, " ", "%20")
    strMsg = strMsg & "Body=" & vstrBodyText
            
    If Len(strMsg) > 2000 Then
       strMsg = Left(strMsg, 2000) & "..."
       If vfTruncateUntilLastVbCrlf = True Then
          For lngIdx = Len(strMsg) To 1 Step -1
             If Mid(strMsg, lngIdx, 6) = "%0d%0a" Then
                strMsg = Left(strMsg, lngIdx + 5) & "..."
                Exit For
             End If
          Next lngIdx
       End If
    End If
    
    lngRet = ShellExecute(0&, vbNullString, strMsg, _
            vbNullString, vbNullString, vbNormalFocus)
    If lngRet >= 42 Then
        Send = RET_OK
    Else
        Send = RET_FAIL
    End If
HandleExit:
    Exit Function
HandleErr:
    Send = Err.Number
    Resume HandleExit
End Function

