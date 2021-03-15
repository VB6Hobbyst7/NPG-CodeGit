<include modSelectionBox.bas>
Dim varArrayList As Variant
Dim varArraySelected As Variant
varArrayList = Array("value1", "value2", "value3")
varArraySelected = SelectionBoxMulti(List:=varArrayList, Prompt:="Select one or more values", SelectionType:=fmMultiSelectMulti)
If Not IsEmpty(varArraySelected) Then 'not cancelled
    <cursor>
End If