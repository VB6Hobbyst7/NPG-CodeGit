'https://msdn.microsoft.com/EN-US/library/office/ff193837.aspx
Dim strMessage As String 
Dim intResponse As Integer 

On Error GoTo ErrorHandler 

' Display the custom dialog box. 
strMessage = "Would you like to delete the current record?" 
intResponse = MsgBox(strMessage, vbYesNo + vbQuestion, "Continue delete?") 

' Check the response. 
If intResponse = vbYes Then 
  Response = 0 'acDataErrContinue 
Else 
  Cancel = True 
End If 

Exit Sub 

ErrorHandler: 
MsgBox "Error #: " & Err.Number & vbCrLf & vbCrLf & Err.Description