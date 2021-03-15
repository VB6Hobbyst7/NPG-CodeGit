VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormSelectionBox 
   Caption         =   "Select Item"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4425
   OleObjectBlob   =   "FormSelectionBox.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormSelectionBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public IsCancelled As Boolean
Public Prompt As String

Public Sub FillList(Values As Variant)
    With Me.ListBox
        Dim iArrayForNext As Long
        .Clear
        For iArrayForNext = LBound(Values) To UBound(Values)
            .AddItem Values(iArrayForNext)
        Next
    End With
End Sub

Private Sub ListBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    btnOk_Click
End Sub

Private Sub UserForm_Initialize()
    IsCancelled = True
    ListBox.SetFocus
End Sub

Private Sub btnCancel_Click()
    Me.Hide
End Sub

Private Sub btnOk_Click()
    If Len(Prompt) > 0 Then
        With Me.ListBox
            If .MultiSelect = fmMultiSelectSingle Then
                If .ListIndex = -1 Then
                    MsgBox Prompt
                    Exit Sub
                End If
            Else
                Dim booHasSelected As Boolean: booHasSelected = False
                For i = 0 To .ListCount - 1
                    If .Selected(i) Then
                        booHasSelected = True
                        Exit For
                    End If
                Next i
                If booHasSelected = False Then
                    MsgBox Prompt
                    Exit Sub
                End If
            End If
        End With
    End If
    IsCancelled = False
    Me.Hide
End Sub
