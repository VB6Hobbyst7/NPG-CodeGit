VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DatePicker 
   ClientHeight    =   1755
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "DatePicker.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DatePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public IsCancelled As Boolean
Public BeginDate As Date
Public EndDate As Date

Private Sub btnCancel_Click()
    Me.Hide
End Sub

Private Sub btnOk_Click()
    Dim EmptyDate As Date
    
    If BeginDate > ComboBox1.value Then
         MsgBox "Date must be after " & BeginDate
         Exit Sub
    End If
    If EndDate > EmptyDate And EndDate < ComboBox1.value Then
         MsgBox "Date must be before " & EndDate
         Exit Sub
    End If

    IsCancelled = False
    Me.Hide
End Sub

Private Sub ComboBox1_DropButtonClick()
    ComboBox1.Enabled = False
    ComboBox1.Enabled = True
    
    Dim frm As New MonthView
    Dim titleBarHeight As Single, borderWidth As Single
    titleBarHeight = Me.Height - Me.InsideHeight - ((Me.Width - Me.InsideWidth) / 2)
    borderWidth = (Me.Width - Me.InsideWidth) / 2
    frm.Top = Me.Top + ComboBox1.Top + ComboBox1.Height + titleBarHeight
    frm.Left = Me.Left + ComboBox1.Left + borderWidth
    Set frm.ParentCombo = ComboBox1
    frm.MonthViewDate = ComboBox1.value
    
    Set frm.Parent = Me
    frm.Show
End Sub

Private Sub UserForm_Initialize()
    IsCancelled = True
    ComboBox1.value = Format(Date, "Short Date")
End Sub
