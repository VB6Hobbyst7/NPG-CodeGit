Dim sep As String: sep = ";"
If Len(<String:List>) = 0 Then
    <String:List> = <String:Item>
Else
    <String:List> = <String:List> & sep & <String:Item>
End If