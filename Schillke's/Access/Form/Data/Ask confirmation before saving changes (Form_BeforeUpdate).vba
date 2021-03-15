'https://msdn.microsoft.com/EN-US/library/office/ff193505.aspx
Dim strMsg As String 
Dim iResponse As Integer 

strMsg = "Do you wish to save the changes?" & Chr(10) 
strMsg = strMsg & "Click Yes to Save or No to Discard changes." 

iResponse = MsgBox(strMsg, vbQuestion + vbYesNo, "Save Record?") 

If iResponse = vbNo Then 

    ' Undo the change. 
    <Application.>DoCmd.RunCommand 292 'acCmdUndo 

    Cancel = True 
End If