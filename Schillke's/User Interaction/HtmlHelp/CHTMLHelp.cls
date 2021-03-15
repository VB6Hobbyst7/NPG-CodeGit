VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CHTMLHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright (C) 2004-2008 CODE-VB, http://www.code-vb.com/, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission.

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Option Compare Text

#Const conAddInCode = False

Private Const mcstrModuleName As String = "CHTMLHelp"


'// Commands to pass to HtmlHelp()
Private Const HH_DISPLAY_TOPIC = &H0
Private Const HH_HELP_FINDER = &H0              '// WinHelp equivalent
Private Const HH_DISPLAY_TOC = &H1
Private Const HH_DISPLAY_INDEX = &H2
Private Const HH_DISPLAY_SEARCH = &H3
Private Const HH_SET_WIN_TYPE = &H4
Private Const HH_GET_WIN_TYPE = &H5
Private Const HH_GET_WIN_HANDLE = &H6
Private Const HH_ENUM_INFO_TYPE = &H7           '// Get Info type name, call repeatedly to enumerate, -1 at end
Private Const HH_SET_INFO_TYPE = &H8            '// Add Info type to filter.
Private Const HH_SYNC = &H9
Private Const HH_RESERVED1 = &HA
Private Const HH_RESERVED2 = &HB
Private Const HH_RESERVED3 = &HC
Private Const HH_KEYWORD_LOOKUP = &HD
Private Const HH_DISPLAY_TEXT_POPUP = &HE       '// display string resource id or text in a popup window
Private Const HH_HELP_CONTEXT = &HF             '// display mapped numeric value in dwData
Private Const HH_TP_HELP_CONTEXTMENU = &H10     '// text popup help, same as WinHelp HELP_CONTEXTMENU
Private Const HH_TP_HELP_WM_HELP = &H11         '// text popup help, same as WinHelp HELP_WM_HELP
Private Const HH_CLOSE_ALL = &H12               '// close all windows opened directly or indirectly by the caller
Private Const HH_ALINK_LOOKUP = &H13            '// ALink version of HH_KEYWORD_LOOKUP
Private Const HH_GET_LAST_ERROR = &H14          '// not currently implemented // See HHERROR.h
Private Const HH_ENUM_CATEGORY = &H15           '// Get category name, call repeatedly to enumerate, -1 at end
Private Const HH_ENUM_CATEGORY_IT = &H16        '// Get category info type members, call repeatedly to enumerate, -1 at end
Private Const HH_RESET_IT_FILTER = &H17         '// Clear the info type filter of all info types.
Private Const HH_SET_INCLUSIVE_FILTER = &H18    '// set inclusive filtering method for untyped topics to be included in display
Private Const HH_SET_EXCLUSIVE_FILTER = &H19    '// set exclusive filtering method for untyped topics to be excluded from display
Private Const HH_INITIALIZE = &H1C              '// Initializes the help system.
Private Const HH_UNINITIALIZE = &H1D            '// Uninitializes the help system.
Private Const HH_PRETRANSLATEMESSAGE = &HFD     '// Pumps messages. (NULL, NULL, MSG*).
Private Const HH_SET_GLOBAL_PROPERTY = &HFC     '// Set a global property. (NULL, NULL, HH_GPROP)

Private Type HH_FTS_QUERY
    cbStruct As Long           '//Specifies the size of the structure.
    fUniCodeStrings As Long    '//TRUE if all strings are Unicode.
    pszSearchQuery As String   '//String containing the search query.
    iProximity As Long         '//Word proximity.
    fStemmedSearch As Long     '//TRUE for StemmedSearch only.
    fTitleOnly As Long         '//TRUE for Title search only.
    fExecute As Long           '//TRUE to initiate the search.
    pszWindow As String        '//Window to display in.
End Type

Private Declare Function HTMLHelp Lib "HHCtrl.ocx" Alias "#14" _
    (ByVal hWndCaller As Long, _
     ByVal pszFile As String, _
     ByVal uCommand As Long, _
     dwData As Long) As Long
Private Declare Function HTMLHelpS Lib "HHCtrl.ocx" Alias "#14" _
    (ByVal hWndCaller As Long, _
     ByVal pszFile As String, _
     ByVal uCommand As Long, _
     dwData As String) As Long
Private Declare Function HTMLHelpSearch Lib "HHCtrl.ocx" Alias "#14" _
    (ByVal hWndCaller As Long, _
     ByVal pszFile As String, _
     ByVal uCommand As Long, _
     dwData As HH_FTS_QUERY) As Long

Private Declare Function IsWindow Lib "user32" (ByVal HWND As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal HWND As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long

Private Const WM_CLOSE = &H10

Private mstrHelpFullPath As String
Private mcolHWNDs As Collection
Private mlngCookie As Long

Private mcolFormsHelpHandlers As Collection

Private mfTerminated As Boolean

Public Sub OpenTOC(ByVal vstrHelpFileNameOrFullPath As String, _
                   Optional ByVal vlngCallerHWND As Long = 0)
' open help TOC
On Error GoTo HandleErr
    OpenTIS vstrHelpFileNameOrFullPath, vlngCallerHWND, HH_DISPLAY_TOC
HandleErr:
End Sub
Public Sub OpenIndex(ByVal vstrHelpFileNameOrFullPath As String, _
                     Optional ByVal vlngCallerHWND As Long = 0, _
                     Optional ByVal vstrIndexKeyWord As String = "")
On Error GoTo HandleErr
' open help Index
    OpenTIS vstrHelpFileNameOrFullPath, vlngCallerHWND, HH_DISPLAY_INDEX, vstrIndexKeyWord
HandleErr:
End Sub
Public Sub OpenSearch(ByVal vstrHelpFileNameOrFullPath As String, _
                      Optional ByVal vlngCallerHWND As Long = 0)
On Error GoTo HandleErr
' open help Search
    OpenTIS vstrHelpFileNameOrFullPath, vlngCallerHWND, HH_DISPLAY_SEARCH
HandleErr:
End Sub

Private Sub OpenTIS(ByVal vstrHelpFileNameOrFullPath As String, _
                   ByVal vlngCallerHWND As Long, _
                   ByVal vlngDisplayOption As Long, _
                   Optional ByVal vstrKeyWord As String = "")
' auxiliary sub used by OpenTOC, OpenIndex and OpenSearch
On Error GoTo HandleErr
Dim lngHelpHWND As Long
    vstrHelpFileNameOrFullPath = HelpFileFullPath(vstrHelpFileNameOrFullPath)
    If vlngDisplayOption = HH_DISPLAY_INDEX Then
        lngHelpHWND = HTMLHelpS(vlngCallerHWND, ByVal vstrHelpFileNameOrFullPath, vlngDisplayOption, ByVal vstrKeyWord)
    ElseIf vlngDisplayOption = HH_DISPLAY_SEARCH Then
        Dim searchIt As HH_FTS_QUERY
        With searchIt
            .cbStruct = Len(searchIt)
            .fUniCodeStrings = 1&
            .pszSearchQuery = "" 'ss: it seems that it doesn't work vstrKeyWord
            .iProximity = 0&
            .fStemmedSearch = 0&
            .fTitleOnly = 0&
            .fExecute = 1&
            .pszWindow = ""
        End With
        lngHelpHWND = HTMLHelpSearch(vlngCallerHWND, ByVal vstrHelpFileNameOrFullPath, vlngDisplayOption, searchIt)
    Else
        lngHelpHWND = HTMLHelp(vlngCallerHWND, ByVal vstrHelpFileNameOrFullPath, vlngDisplayOption, ByVal 0)
    End If
    If lngHelpHWND <> 0 Then AddHWND vlngCallerHWND, lngHelpHWND
HandleErr:
End Sub

Public Sub OpenTopic(ByVal vstrTopicFileName As String, _
                     ByVal vstrHelpFileNameOrFullPath As String, _
                     Optional ByVal vlngCallerHWND As Long = 0)
' open help topic
On Error GoTo HandleErr
Dim lngHelpHWND As Long
    vstrHelpFileNameOrFullPath = HelpFileFullPath(vstrHelpFileNameOrFullPath)
    lngHelpHWND = HTMLHelpS(vlngCallerHWND, ByVal vstrHelpFileNameOrFullPath, HH_DISPLAY_TOPIC, ByVal vstrTopicFileName)
    If lngHelpHWND <> 0 Then AddHWND vlngCallerHWND, lngHelpHWND
HandleErr:
End Sub

Public Sub OpenContext(ByVal vlngHelpContextId As Long, _
                       ByVal vstrHelpFileNameOrFullPath As String, _
                       Optional ByVal vlngCallerHWND As Long = 0)
' open help context
On Error GoTo HandleErr
Dim lngHelpHWND As Long
    vstrHelpFileNameOrFullPath = HelpFileFullPath(vstrHelpFileNameOrFullPath)
    lngHelpHWND = HTMLHelp(vlngCallerHWND, ByVal vstrHelpFileNameOrFullPath, HH_HELP_CONTEXT, ByVal vlngHelpContextId)
    If lngHelpHWND <> 0 Then AddHWND vlngCallerHWND, lngHelpHWND
HandleErr:
End Sub

Public Sub CloseHelpWindows(Optional ByVal vlngCallerHWND As Long = 0)
' close help windows
On Error GoTo HandleErr
Dim lngIdx As Long
Dim lngCallerHWND As Long
Dim lngHelpHWND As Long
Dim strHWNDs As String
    If mcolHWNDs.Count > 0 Then
        For lngIdx = mcolHWNDs.Count To 1 Step -1
           strHWNDs = CStr(mcolHWNDs(lngIdx))
           lngCallerHWND = Val(strHWNDs)
           lngHelpHWND = CLng(Mid(strHWNDs, InStr(strHWNDs, "_") + 1))
           If (lngCallerHWND = vlngCallerHWND) Or (vlngCallerHWND = 0) Then
              If IsWindow(lngHelpHWND) <> 0 Then
                 ' sendmessage is a quick compare to HTML Help API, which delays 0.5s
                 SendMessage lngHelpHWND, WM_CLOSE, 0, 0
                 DoEvents
              End If
              mcolHWNDs.Remove lngIdx
           End If
        Next lngIdx
    End If
HandleErr:
End Sub

Public Sub CloseAll()
' close all help windows
On Error GoTo HandleErr
    CloseHelpWindows 0
HandleErr:
End Sub

Private Sub AddHWND(ByVal vlngCallerHWND As Long, ByVal vlngHelpHWND As Long)
    If vlngHelpHWND > 0 Then
        mcolHWNDs.Add CStr(vlngCallerHWND) & "_" & CStr(vlngHelpHWND)
    End If
End Sub

Private Sub Class_Initialize()
    Set mcolHWNDs = New Collection
    #If conAddInCode Then
    ' may give trouble (gpf) if used in Add-ins  comment it by CC
    #Else
    ' must have when used in standalone application
    HTMLHelp 0, 0, HH_INITIALIZE, mlngCookie  ' // Cookie returned by Hhctrl.ocx.
    #End If
    Set mcolFormsHelpHandlers = New Collection
End Sub

Private Sub Class_Terminate()
    If mfTerminated = False Then
        CloseAll
        #If conAddInCode Then
        ' may give trouble (gpf) when used in add-in
        #Else
        ' must have when used in standalone app
        If mlngCookie <> 0 Then
           HTMLHelp 0, 0, HH_UNINITIALIZE, ByVal mlngCookie
        End If
        #End If
        mfTerminated = True
        Set mcolHWNDs = Nothing
        MsgBox mcstrModuleName & " terminated."
    End If
End Sub

'Public Sub InitFormHelpHandler( _
'             ByRef rfrm As Access.Form, _
'             ByRef rcmdHelp As Access.CommandButton,
'Public Function InitFormHelpHandler( _
'             ByRef rfrm As Object, _
'             ByRef rcmdHelp As MSForms.CommandButton, _
'             Optional ByVal vfKeyPreview As Boolean = True) As CMSFormHtmlHelpHandler
'' Init form help handler
'On Error GoTo HandleErr
''Dim obj As CFormHtmlHelpHandler
''    Set obj = New CFormHtmlHelpHandler
'Dim obj As CMSFormHtmlHelpHandler
'    Set obj = New CMSFormHtmlHelpHandler
'    obj.Init rfrm, rfrm.HWND, rcmdHelp, robjParent:=Me, vfKeyPreview:=vfKeyPreview
'    mcolFormsHelpHandlers.Add obj, CStr(rfrm.HWND)
'    Set InitFormHelpHandler = obj
'HandleExit:
'    Exit Function
'HandleErr:
'    ErrorHandle Err, Erl, , "InitFormHelpHandler", mcstrModuleName
'    Resume HandleExit
'    Resume
'End Function
'
''Public Sub FormClosing(ByRef rfrm As Access.Form)
'Public Sub FormClosing(ByRef rfrm As Object)
'' remove form help handler on form Close event
'On Error GoTo HandleErr
'    mcolFormsHelpHandlers.Remove CStr(rfrm.HWND)
'HandleErr:
'End Sub
'
'Public Property Get Handle() As Long
'Static lngHandle As Long
'    lngHandle = lngHandle + 1
'    Handle = lngHandle
'End Property

