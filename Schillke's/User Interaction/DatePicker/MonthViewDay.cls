VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MonthViewDay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents DayLabel As MSForms.Label
Attribute DayLabel.VB_VarHelpID = -1
Private Sub DayLabel_Click()
    DayLabel.Parent.Parent.Controls("MyDate").Text = DayLabel.Tag
End Sub

Private Sub DayLabel_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    DayLabel.Font.Bold = True
End Sub

Private Sub DayLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim i As Integer
    For i = 1 To 42
        If DayLabel.Parent.Controls("Label" & i).ForeColor = vbBlue Then
            DayLabel.Parent.Controls("Label" & i).ForeColor = vbBlack
        End If
    Next
    If DayLabel.ForeColor = vbBlack Then
        DayLabel.ForeColor = vbBlue
        DayLabel.Parent.Tag = DayLabel.Name
    End If
End Sub

