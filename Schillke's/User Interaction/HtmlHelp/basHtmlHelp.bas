Attribute VB_Name = "basHtmlHelp"
'<include CHTMLHelp.cls>
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright (C) 2004-2016 CODE VBA, http://codevba.com, http://helpgenerator.com All Rights Reserved.
' Some pages may also contain other copyrights by the author.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Distribution: You can freely use this code in your own applications,
' but you may not reproduce or publish this code on any web site, online service,
' or distribute as source on any media without express permission.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' HTML HELP singleton object running in the same thread as the host application
Private mobjHTMLHelp As CHTMLHelp

'TODO:
'- specify cstHelpFile: "MyHelpFile.chm" or "C:\..\MyHelpFile.chm"
Const cstHelpFile As String = ""

Public Function AppPath() As String
Dim strAppPath As String
'    strAppPath = Excel.Application.ActiveWorkbook.Path & "\"
'    strAppPath = Access.Application.CurrentProject.Path & "\"
'    ... or other. Often it is useful to have the path determined relative to the item being called from
    If Len(strAppPath) = 0 Then MsgBox Prompt:="Function AppPath must be implemented", Buttons:=vbOKOnly + vbInformation
    AppPath = strAppPath
End Function

Public Function HelpFile() As String
    Dim strHelpFile As String
    'strHelpFile = Excel.Application.ActiveWorkbook.VBProject.HelpFile
    'strHelpFile = Access.Screen.ActiveForm.HelpFile
    If Len(strHelpFile) = 0 Then strHelpFile = cstHelpFile
    If Len(strHelpFile) = 0 Then MsgBox Prompt:="Const cstHelpFile must be specified", Buttons:=vbOKOnly + vbInformation
    HelpFile = strHelpFile
End Function

Public Property Get HTMLHelp() As CHTMLHelp
' Get HTML Help singleton object instance - custom class clsHtmlHelp
    If mobjHTMLHelp Is Nothing Then
       Set mobjHTMLHelp = New CHTMLHelp
    End If
    Set HTMLHelp = mobjHTMLHelp
End Property

Public Function HelpFileFullPath(ByVal vstrHelpFileNameOrFullPath As String) As String
    'Allow the calling procedure to specify a file
    If Len(vstrHelpFileNameOrFullPath) = 0 Then
        vstrHelpFileNameOrFullPath = HelpFile()
    End If
    
    If (InStr(vstrHelpFileNameOrFullPath, "\") > 0 Or InStr(vstrHelpFileNameOrFullPath, "/") > 0) Then
        'vstrHelpFileNameOrFullPath includes Path
        HelpFileFullPath = vstrHelpFileNameOrFullPath
    Else
        HelpFileFullPath = AppPath() & vstrHelpFileNameOrFullPath
    End If
End Function

Public Property Get helpContextId() As String
On Error Resume Next
    helpContextId = ActiveWorkbook.VBProject.helpContextId '10
End Property

Public Sub HtmlHelpOpenTOC()
' open help TOC
    HTMLHelp.OpenTOC HelpFile
End Sub

Public Sub HtmlHelpOpenIndex()
' open help Index
    Dim strHelpIndexKeyword As String
    strHelpIndexKeyword = "UserForm1"
    HTMLHelp.OpenIndex HelpFile, vstrIndexKeyWord:=strHelpIndexKeyword
End Sub

Public Sub HtmlHelpOpenSearch()
' open help Search
    HTMLHelp.OpenSearch HelpFile
End Sub

Public Sub HtmlHelpOpenTopicByContextId1()
' open help topic using HelpContextId
    HTMLHelp.OpenContext helpContextId, HelpFile
End Sub

Public Sub HtmlHelpOpenTopicByBoomark()
' open help bookmark
Dim strHtmlFileName As String
Dim strHtmlFileBookmark As String
Dim strBookmark As String
    strHtmlFileName = "XLHelpSample1_Worksheet_Customers_Qty.htm"
    strHtmlFileBookmark = "B10"
    
    strBookmark = strHtmlFileName
    If Len(strHtmlFileBookmark) > 0 Then
       strBookmark = strBookmark & "#" & strHtmlFileBookmark
    End If
    HTMLHelp.OpenTopic strBookmark, HelpFile
End Sub



