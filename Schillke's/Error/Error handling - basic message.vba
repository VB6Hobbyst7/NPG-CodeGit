<NoNewLine>
On Error GoTo HandleError
<cursor>
<top>
<bottom>
HandleExit:
Exit <proceduretype>
HandleError:
MsgBox Err.Description
Resume HandleExit