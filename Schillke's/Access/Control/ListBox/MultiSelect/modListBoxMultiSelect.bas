Attribute VB_Name = "modListBoxMultiSelect"
Public Sub ClearList(lst As ListBox)
'Precondition: MultiSelect > 0
'None: 0 (Default) Multiple selection isn't allowed.
'Simple: 1
'Extended: 2
    Dim varItem As Variant

    For Each varItem In lst.ItemsSelected
        lst.Selected(varItem) = False
    Next
End Sub

Public Sub SelectAll(lst As ListBox)
'Precondition: MultiSelect > 0
'None: 0 (Default) Multiple selection isn't allowed.
'Simple: 1
'Extended: 2
    Dim lngRow As Long

    For lngRow = 0 To lst.ListCount - 1
        lst.Selected(lngRow) = True
    Next
End Sub
