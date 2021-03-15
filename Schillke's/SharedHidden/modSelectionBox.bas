Attribute VB_Name = "modSelectionBox"
'<include FormSelectionBox.frm>
Option Explicit

Public Function SelectionBoxSingle(List As Variant, Optional Prompt As String, Optional Title As String) As String
Dim varReturnList As Variant
    varReturnList = SelectionBoxMulti(List:=List, Prompt:=Prompt, SelectionType:=fmMultiSelectSingle, Title:=Title)
    
    If Not IsEmpty(varReturnList) Then
        If Not Len(Join(varReturnList, "")) = 0 Then
            SelectionBoxSingle = varReturnList(UBound(varReturnList))
        End If
    End If
    
End Function

Public Function SelectionBoxMulti(List As Variant, Optional Prompt As String, Optional SelectionType As MSForms.fmMultiSelect = fmMultiSelectMulti, Optional Title As String) As Variant
Dim SelectedValues() As Variant
    With FormSelectionBox
        If Len(Title) > 0 Then .Caption = Title
        
        .FillList List
        .Prompt = Prompt
        .ListBox.MultiSelect = SelectionType
        .Show
               
        If Not .IsCancelled Then
        Dim i As Integer, j As Integer
            With .ListBox
                For i = 0 To .ListCount - 1
                    If .Selected(i) Then
                        ReDim Preserve SelectedValues(j)
                        SelectedValues(j) = .List(i)
                        j = j + 1
                    End If
                Next i
            End With
            SelectionBoxMulti = SelectedValues
        End If
    End With
    Unload FormSelectionBox
End Function


