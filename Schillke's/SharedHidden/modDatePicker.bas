Attribute VB_Name = "modDatePicker"
'<include FormDatePicker.frm>
Public Function PickDate(Optional Default As Date, Optional BeginDate As Date, Optional EndDate As Date, _
                        Optional Prompt As String, Optional Title As String) As Variant
    With FormDatePicker
        If Len(Title) > 0 Then
            .Caption = Title
        Else
            .Caption = Application.Name
        End If
        
        If Len(Prompt) > 0 Then !Label1.Caption = Prompt
        
        Dim EmptyDate As Date
        If Default = EmptyDate Then Default = Date
        !DTPicker1.Value = Default
        If BeginDate > EmptyDate Then .BeginDate = BeginDate
        If EndDate > EmptyDate Then .EndDate = EndDate
        
        .Show vbModal
    
        If Not .IsCancelled Then
            PickDate = !DTPicker1.Value
        End If
    End With
    
    Unload FormDatePicker
End Function

