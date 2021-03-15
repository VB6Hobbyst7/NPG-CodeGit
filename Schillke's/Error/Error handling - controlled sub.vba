<NoNewLine>
On Error GoTo HandleError
<cursor>
<top>
<bottom>
Exit <proceduretype>
HandleError:
Err.Raise Err.Number, IIf(Err.Source = Application.VBE.ActiveVBProject.Name, "<modulename>.<procedurename>" & " " & Erl, Err.Source), Err.Description