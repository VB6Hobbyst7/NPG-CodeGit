Dim strMsg As String: strMsg = "Do you want to add " & NewData & " to the table?"
If MsgBox(Prompt:=strMsg, Buttons:=vbQuestion + vbYesNo, Title:="Add new name?") = vbNo Then
    Response = 0 'acDataErrContinue
Else
    Set <Database> = <Application.>CurrentDb
    Set <Recordset> = <Database>.OpenRecordset("<cursor>", dbOpenDynaset) 'TODO: Add table name
    On Error Resume Next
    <Recordset>.AddNew
        <Recordset>! = NewData 'TODO: Add field name after rst!
    <Recordset>.Update    
    If Err Then
        MsgBox "An error occurred. Please try again."
        Response = 0 'acDataErrContinue
    Else
        Response = 2 'acDataErrAdded
    End If
End If
