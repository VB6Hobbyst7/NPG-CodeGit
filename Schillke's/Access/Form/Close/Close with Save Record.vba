<NoNewLine>
On Error GoTo HandleError
<Application.>DoCmd.RunCommand 97 'acCmdSaveRecord = 97
<Application.>DoCmd.Close ObjectType:=2, ObjectName:=Me.Name  'acForm = 2
'<Application.>DoCmd.Close 2, Me.Name  'acForm = 2
HandleExit:
Exit Sub
HandleError:
MsgBox Err.Description
Resume HandleExit