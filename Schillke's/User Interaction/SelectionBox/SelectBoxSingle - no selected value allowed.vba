<include modSelectionBox.bas>
Dim varArrayList As Variant
Dim strSelected As String
varArrayList = Array("value1", "value2", "value3")
strSelected = SelectionBoxSingle(List:=varArrayList)
If Len(strSelected) > 0 Then
    
End If