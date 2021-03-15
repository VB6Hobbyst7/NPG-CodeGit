VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MonthView 
   Caption         =   "MonthView"
   ClientHeight    =   2970
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2325
   OleObjectBlob   =   "MonthView.frx":0000
End
Attribute VB_Name = "MonthView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#If Win64 Then

Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" ( _
    ByVal lpClassName As String, _
    ByVal lpWindowName As String) As LongPtr

Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongPtrA" ( _
    ByVal hWnd As LongPtr, _
    ByVal nIndex As Long) As LongPtr
 
Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongPtrA" ( _
    ByVal hWnd As LongPtr, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As LongPtr) As LongPtr

Private Declare PtrSafe Function DrawMenuBar Lib "user32" ( _
    ByVal hWnd As LongPtr) As Long

#Else

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
    ByVal lpClassName As String, _
    ByVal lpWindowName As String) As Long

Private Declare Function GetWindowLong Lib "user32" _
    Alias "GetWindowLongA" ( _
    ByVal hWnd As Long, _
    ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
    ByVal hWnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long

Private Declare Function DrawMenuBar Lib "user32" ( _
    ByVal hWnd As Long) As Long

#End If


Private mMonthViewDate As Date

Public Labels As New Collection
Public BlueLabel As MSForms.Label
Public ParentCombo As MSForms.ComboBox
Public Parent As MSForms.UserForm

Public Property Let MonthViewDate(value As Date)
    mMonthViewDate = value
    If Not ParentCombo Is Nothing Then
        ParentCombo.value = value
    End If
    Form_Draw
End Property

Public Property Get MonthViewDate() As Date
    MonthViewDate = mMonthViewDate
End Property

Sub RemoveTitleBar(frm As Object)
    
    Dim hMenu           As Long
    #If Win64 Then
        Dim mhWndForm       As LongPtr
        Dim lStyle          As LongPtr
    #Else
        Dim mhWndForm       As Long
        Dim lStyle          As Long
    #End If
         
    If Val(Application.Version) < 9 Then
        mhWndForm = FindWindow("ThunderXFrame", frm.Caption) 'for Office 97 version
    Else
        mhWndForm = FindWindow("ThunderDFrame", frm.Caption) 'for office 2000 or above
    End If
    lStyle = GetWindowLong(mhWndForm, -16)
    lStyle = lStyle And Not &HC00000
    SetWindowLong mhWndForm, -16, lStyle
    DrawMenuBar mhWndForm
End Sub

Private Sub btnDummy_Click()
    Unload Me
End Sub

Private Sub btnMonthBack_Click()
    ChangeDate "m", -1
End Sub

Private Sub btnMonthForward_Click()
    ChangeDate "m", 1
End Sub

Private Sub btnYearBack_Click()
    ChangeDate "yyyy", -1
End Sub

Private Sub btnYearForward_Click()
    ChangeDate "yyyy", 1
End Sub

Private Sub FrameBottom_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RemoveHighlight
End Sub

Private Sub FrameTop_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RemoveHighlight
End Sub

Private Sub RemoveHighlight()
If Not FrameBottom.Tag = "" Then
    Me.Controls(FrameBottom.Tag).ForeColor = vbBlack
    FrameBottom.Tag = ""
End If
End Sub

Private Sub LabelToday_Click()
    Me.MonthViewDate = VBA.Date
    Form_Draw
End Sub

Private Sub MyDate_Change()
    MonthViewDate = MyDate
    Unload Me
End Sub

Private Sub UserForm_Click()
    Unload Me
End Sub

Private Sub UserForm_Deactivate()
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    RemoveTitleBar Me
    Me("LabelToday").Caption = "Today: " & Format(Date, "dd.mm.yyyy")
    MonthViewDate = VBA.Date
    Dim i As Integer
    For i = 1 To 42
        Labels.Add New MonthViewDay, "Label" & i
        Set Labels("Label" & i).DayLabel = Me("Label" & i)
        If i < 8 Then Me("LabelWeekDay" & i).Caption = Left(WeekdayName(i, True, 2), 2)
    Next
    Form_Draw
End Sub

Private Sub Form_Draw()
    Dim sCtl As MSForms.Label
    Dim sDtmDate As Date
    Dim sDtmStartDate As Date
    Dim sIntCounter As Integer
    Dim sIntMonth As Integer

    Me("cap").Caption = Format(MonthViewDate, "mmmm yyyy")
    
    sIntMonth = Month(MonthViewDate)
    sDtmStartDate = DateSerial(Year(MonthViewDate), Month(MonthViewDate), 1)
    sDtmStartDate = sDtmStartDate - Weekday(sDtmStartDate, vbUseSystem) + 1
   
    For sIntCounter = 0 To 41
        sDtmDate = sDtmStartDate + sIntCounter
        Set sCtl = Me.Controls("Label" & CStr(sIntCounter + 1))
        sCtl.TextAlign = fmTextAlignCenter
        With sCtl
            .Caption = Format(sDtmDate, "d")
            .Tag = sDtmDate
        
            Select Case sDtmDate
                Case VBA.Date
                    .ForeColor = vbRed
                Case Else
                    If Month(sDtmDate) = sIntMonth Then
                        .ForeColor = vbBlack
                    Else
                        .ForeColor = vbGrayText
                    End If
            End Select
            If sDtmDate = mMonthViewDate Then
                .BackColor = vbYellow
            Else
                .BackColor = vbWhite
            End If
    
            .FontBold = (sDtmDate = Date)
        End With
    
    Next sIntCounter
    Me.Repaint
End Sub

Public Sub ChangeDate(ByVal Interval As String, ByVal Direction As Long)
    On Error Resume Next
    MonthViewDate = DateAdd(Interval, Direction, MonthViewDate)
    Form_Draw
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
  Select Case KeyCode
    Case vbKeyLeft
        ChangeDate "d", -1
    Case vbKeyRight
        ChangeDate "d", 1
    Case vbKeyUp
        ChangeDate "d", -7
    Case vbKeyDown
        ChangeDate "d", 7
    Case Else
        Unload Me
        KeyCode = 0
    End Select
End Sub

Private Sub UserForm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Debug.Print "!"
End Sub

Private Sub UserForm_Terminate()
'    EnableWindow Application.hWnd, Modal
''    Parent.Show 1
    Set Parent = Nothing
    Set ParentCombo = Nothing
    'ReleaseCapture '++
End Sub


