<include modSelectionBox.bas>
Dim varArrayList As Variant
Dim varArraySelected As Variant
varArrayList = Array("value1", "value2", "value3")
varArraySelected = SelectionBoxMulti(List:=varArrayList, SelectionType:=fmMultiSelectMulti)
If Not IsEmpty(varArraySelected) Then 'not cancelled
    If Len(Join(varArraySelected, "")) > 0 Then 'at least one item selected
        <cursor>
    End If
End If