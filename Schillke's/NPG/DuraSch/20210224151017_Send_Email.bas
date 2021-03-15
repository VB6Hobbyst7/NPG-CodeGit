Attribute VB_Name = "Send_Email"
Sub EmailGenerator()

On Error GoTo TheEnd

'Insertion of Message Box to insure correct email recipient.  If yes, then run program.

Cells(2, 8).Select
Emailname = Cells(2, 9).Value
Question = "Confirm you are:  " & Emailname
MSG1 = MsgBox(Question, vbOKCancel, "Confirm Email Recipient (NOTE: Should be to yourself!)")

If MSG1 = vbOK Then

Dim Email_Subject, Email_Send_From, Email_Send_To, _
 Email_Cc, Email_Bcc, Email_Body As String
Dim Mail_Object, Mail_Single As Variant

Dim VSheetName As String, DriverIndicator As String

Dim FTVehicleID, FTDriverID, FTStatusID, FTString
Dim VehicleID, DriverID, StatusID, NFTString
Dim WOVehicleID, WODriverID, WOName, WOStatus, WODesc, WOString

FTVehicleID = Array(1, 2, 3, 4, 5, 6, 7, 8, 9)
FTDriverID = Array(1, 2, 3, 4, 5, 6, 7, 8, 9)
FTStatusID = Array(1, 2, 3, 4, 5, 6, 7, 8, 9)
FTMilesID = Array(1, 2, 3, 4, 5, 6, 7, 8, 9)
FTString = Array(1, 2, 3, 4, 5, 6, 7, 8, 9)

VehicleID = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19)
DriverID = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19)
StatusID = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19)
MilesID = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19)
NFTString = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19)

WOVehicleID = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24)
WODriverID = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24)
WOName = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24)
WOStatus = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24)
WODesc = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24)
WOString = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24)

'This section looks for actual vehicle runners in the Field Test & Non-Field Test Groups BASED ON STATUS not driver
      x = 1 ' Field Test
For I = 5 To 12
  'Sheets(SheetName).Select
  DriverIndicator = Cells(I, 5).Value
    If DriverIndicator <> "" Then
       FTVehicleID(x) = Cells(I, 2).Value
       FTDriverID(x) = Cells(I, 5).Value
       FTMilesID(x) = Cells(I, 7).Value
       FTStatusID(x) = Cells(I, 8).Value
      x = x + 1
    End If
Next I
x = x - 1

      y = 1 ' Non-Field Test
For I = 14 To 31
  'Sheets(SheetName).Select
  DriverIndicator = Cells(I, 5).Value
    If DriverIndicator <> "" Then
       VehicleID(y) = Cells(I, 2).Value
       DriverID(y) = Cells(I, 5).Value
       MilesID(y) = Cells(I, 7).Value
       StatusID(y) = Cells(I, 8).Value
      y = y + 1
    End If
Next I
y = y - 1

      Z = 1 ' Work Orders
For I = 33 To 55
  'Sheets(SheetName).Select
  DriverIndicator = Cells(I, 2).Value 'This section looks for Work Orders BASED ON Vehicle ID not Status
    If DriverIndicator <> "" Then
       WOVehicleID(Z) = Cells(I, 2).Value
       WODesc(Z) = Cells(I, 4).Value
       WODriverID(Z) = Cells(I, 5).Value
       WOStatus(Z) = Cells(I, 8).Value
       WOName(Z) = Cells(I, 9).Value
      Z = Z + 1
    End If
Next I
Z = Z - 1

' Strings together Field Test & Non-Field Test Info to put in Message Body
    For a = 1 To x ' Field Test
        FTString(a) = Chr(149) & "   " & FTVehicleID(a) & " (" & FTDriverID(a) & ")     Status:  " & FTStatusID(a) & " at End of shift (" & FTMilesID(a) & " miles)" & Chr(13) & Chr(10)
    Next a

If x = 0 Then
FTString = ""
ElseIf x = 1 Then
FTString = FTString(1)
ElseIf x = 2 Then
FTString = FTString(1) & FTString(2)
ElseIf x = 3 Then
FTString = FTString(1) & FTString(2) & FTString(3)
ElseIf x = 4 Then
FTString = FTString(1) & FTString(2) & FTString(3) & FTString(4)
ElseIf x = 5 Then
FTString = FTString(1) & FTString(2) & FTString(3) & FTString(4) & FTString(5)
ElseIf x = 6 Then
FTString = FTString(1) & FTString(2) & FTString(3) & FTString(4) & FTString(5) & FTString(6)
ElseIf x = 7 Then
FTString = FTString(1) & FTString(2) & FTString(3) & FTString(4) & FTString(5) & FTString(6) & FTString(7)
ElseIf x = 8 Then
FTString = FTString(1) & FTString(2) & FTString(3) & FTString(4) & FTString(5) & FTString(6) & FTString(7) & FTString(8)
End If
    
    For b = 1 To y ' Non-Field Test
        NFTString(b) = Chr(149) & "   " & VehicleID(b) & " (" & DriverID(b) & ")     Status:  " & StatusID(b) & " at End of shift (" & MilesID(b) & " miles)" & Chr(13) & Chr(10)
    Next b

If y = 0 Then
NFTString = ""
ElseIf y = 1 Then
NFTString = NFTString(1)
ElseIf y = 2 Then
NFTString = NFTString(1) & NFTString(2)
ElseIf y = 3 Then
NFTString = NFTString(1) & NFTString(2) & NFTString(3)
ElseIf y = 4 Then
NFTString = NFTString(1) & NFTString(2) & NFTString(3) & NFTString(4)
ElseIf y = 5 Then
NFTString = NFTString(1) & NFTString(2) & NFTString(3) & NFTString(4) & NFTString(5)
ElseIf y = 6 Then
NFTString = NFTString(1) & NFTString(2) & NFTString(3) & NFTString(4) & NFTString(5) & NFTString(6)
ElseIf y = 7 Then
NFTString = NFTString(1) & NFTString(2) & NFTString(3) & NFTString(4) & NFTString(5) & NFTString(6) & _
NFTString(7)
ElseIf y = 8 Then
NFTString = NFTString(1) & NFTString(2) & NFTString(3) & NFTString(4) & NFTString(5) & NFTString(6) & _
NFTString(7) & NFTString(8)
ElseIf y = 9 Then
NFTString = NFTString(1) & NFTString(2) & NFTString(3) & NFTString(4) & NFTString(5) & NFTString(6) & _
NFTString(7) & NFTString(8) & NFTString(9)
ElseIf y = 10 Then
NFTString = NFTString(1) & NFTString(2) & NFTString(3) & NFTString(4) & NFTString(5) & NFTString(6) & _
NFTString(7) & NFTString(8) & NFTString(9) & NFTString(10)
ElseIf y = 11 Then
NFTString = NFTString(1) & NFTString(2) & NFTString(3) & NFTString(4) & NFTString(5) & NFTString(6) & _
NFTString(7) & NFTString(8) & NFTString(9) & NFTString(10) & NFTString(11)
ElseIf y = 12 Then
NFTString = NFTString(1) & NFTString(2) & NFTString(3) & NFTString(4) & NFTString(5) & NFTString(6) & _
NFTString(7) & NFTString(8) & NFTString(9) & NFTString(10) & NFTString(11) & NFTString(12)
ElseIf y = 13 Then
NFTString = NFTString(1) & NFTString(2) & NFTString(3) & NFTString(4) & NFTString(5) & NFTString(6) & _
NFTString(7) & NFTString(8) & NFTString(9) & NFTString(10) & NFTString(11) & NFTString(12) & NFTString(13)
ElseIf y = 14 Then
NFTString = NFTString(1) & NFTString(2) & NFTString(3) & NFTString(4) & NFTString(5) & NFTString(6) & _
NFTString(7) & NFTString(8) & NFTString(9) & NFTString(10) & NFTString(11) & NFTString(12) & NFTString(13) & _
NFTString(14)
ElseIf y = 15 Then
NFTString = NFTString(1) & NFTString(2) & NFTString(3) & NFTString(4) & NFTString(5) & NFTString(6) & _
NFTString(7) & NFTString(8) & NFTString(9) & NFTString(10) & NFTString(11) & NFTString(12) & NFTString(13) & _
NFTString(14) & NFTString(15)
ElseIf y = 16 Then
NFTString = NFTString(1) & NFTString(2) & NFTString(3) & NFTString(4) & NFTString(5) & NFTString(6) & _
NFTString(7) & NFTString(8) & NFTString(9) & NFTString(10) & NFTString(11) & NFTString(12) & NFTString(13) & _
NFTString(14) & NFTString(15) & NFTString(16)
ElseIf y = 17 Then
NFTString = NFTString(1) & NFTString(2) & NFTString(3) & NFTString(4) & NFTString(5) & NFTString(6) & _
NFTString(7) & NFTString(8) & NFTString(9) & NFTString(10) & NFTString(11) & NFTString(12) & NFTString(13) & _
NFTString(14) & NFTString(15) & NFTString(16) & NFTString(17)
ElseIf y = 18 Then
NFTString = NFTString(1) & NFTString(2) & NFTString(3) & NFTString(4) & NFTString(5) & NFTString(6) & _
NFTString(7) & NFTString(8) & NFTString(9) & NFTString(10) & NFTString(11) & NFTString(12) & NFTString(13) & _
NFTString(14) & NFTString(15) & NFTString(16) & NFTString(17) & NFTString(18)
End If
    
    For c = 1 To Z ' Work Orders
        WOString(c) = Chr(149) & "   " & WOName(c) & ":  " & WOVehicleID(c) & ":  " & WODesc(c) & " (" & WODriverID(c) & ")     Status:  " & WOStatus(c) & Chr(13) & Chr(10)
    Next c

If Z = 0 Then
WOString = ""
ElseIf Z = 1 Then
WOString = WOString(1)
ElseIf Z = 2 Then
WOString = WOString(1) & WOString(2)
ElseIf Z = 3 Then
WOString = WOString(1) & WOString(2) & WOString(3)
ElseIf Z = 4 Then
WOString = WOString(1) & WOString(2) & WOString(3) & WOString(4)
ElseIf Z = 5 Then
WOString = WOString(1) & WOString(2) & WOString(3) & WOString(4) & WOString(5)
ElseIf Z = 6 Then
WOString = WOString(1) & WOString(2) & WOString(3) & WOString(4) & WOString(5) & WOString(6)
ElseIf Z = 7 Then
WOString = WOString(1) & WOString(2) & WOString(3) & WOString(4) & WOString(5) & WOString(6) & _
WOString(7)
ElseIf Z = 8 Then
WOString = WOString(1) & WOString(2) & WOString(3) & WOString(4) & WOString(5) & WOString(6) & _
WOString(7) & WOString(8)
ElseIf Z = 9 Then
WOString = WOString(1) & WOString(2) & WOString(3) & WOString(4) & WOString(5) & WOString(6) & _
WOString(7) & WOString(8) & WOString(9)
ElseIf Z = 10 Then
WOString = WOString(1) & WOString(2) & WOString(3) & WOString(4) & WOString(5) & WOString(6) & _
WOString(7) & WOString(8) & WOString(9) & WOString(10)
ElseIf Z = 11 Then
WOString = WOString(1) & WOString(2) & WOString(3) & WOString(4) & WOString(5) & WOString(6) & _
WOString(7) & WOString(8) & WOString(9) & WOString(10) & WOString(11)
ElseIf Z = 12 Then
WOString = WOString(1) & WOString(2) & WOString(3) & WOString(4) & WOString(5) & WOString(6) & _
WOString(7) & WOString(8) & WOString(9) & WOString(10) & WOString(11) & WOString(12)
ElseIf Z = 13 Then
WOString = WOString(1) & WOString(2) & WOString(3) & WOString(4) & WOString(5) & WOString(6) & _
WOString(7) & WOString(8) & WOString(9) & WOString(10) & WOString(11) & WOString(12) & WOString(13)
ElseIf Z = 14 Then
WOString = WOString(1) & WOString(2) & WOString(3) & WOString(4) & WOString(5) & WOString(6) & _
WOString(7) & WOString(8) & WOString(9) & WOString(10) & WOString(11) & WOString(12) & WOString(13) & _
WOString(14)
ElseIf Z = 15 Then
WOString = WOString(1) & WOString(2) & WOString(3) & WOString(4) & WOString(5) & WOString(6) & _
WOString(7) & WOString(8) & WOString(9) & WOString(10) & WOString(11) & WOString(12) & WOString(13) & _
WOString(14) & WOString(15)
ElseIf Z = 16 Then
WOString = WOString(1) & WOString(2) & WOString(3) & WOString(4) & WOString(5) & WOString(6) & _
WOString(7) & WOString(8) & WOString(9) & WOString(10) & WOString(11) & WOString(12) & WOString(13) & _
WOString(14) & WOString(15) & WOString(16)
ElseIf Z = 17 Then
WOString = WOString(1) & WOString(2) & WOString(3) & WOString(4) & WOString(5) & WOString(6) & _
WOString(7) & WOString(8) & WOString(9) & WOString(10) & WOString(11) & WOString(12) & WOString(13) & _
WOString(14) & WOString(15) & WOString(16) & WOString(17)
ElseIf Z = 18 Then
WOString = WOString(1) & WOString(2) & WOString(3) & WOString(4) & WOString(5) & WOString(6) & _
WOString(7) & WOString(8) & WOString(9) & WOString(10) & WOString(11) & WOString(12) & WOString(13) & _
WOString(14) & WOString(15) & WOString(16) & WOString(17) & WOString(18)
ElseIf Z = 19 Then
WOString = WOString(1) & WOString(2) & WOString(3) & WOString(4) & WOString(5) & WOString(6) & _
WOString(7) & WOString(8) & WOString(9) & WOString(10) & WOString(11) & WOString(12) & WOString(13) & _
WOString(14) & WOString(15) & WOString(16) & WOString(17) & WOString(18) & WOString(19)
ElseIf Z = 20 Then
WOString = WOString(1) & WOString(2) & WOString(3) & WOString(4) & WOString(5) & WOString(6) & _
WOString(7) & WOString(8) & WOString(9) & WOString(10) & WOString(11) & WOString(12) & WOString(13) & _
WOString(14) & WOString(15) & WOString(16) & WOString(17) & WOString(18) & WOString(19) & _
WOString(20)
ElseIf Z = 21 Then
WOString = WOString(1) & WOString(2) & WOString(3) & WOString(4) & WOString(5) & WOString(6) & _
WOString(7) & WOString(8) & WOString(9) & WOString(10) & WOString(11) & WOString(12) & WOString(13) & _
WOString(14) & WOString(15) & WOString(16) & WOString(17) & WOString(18) & WOString(19) & _
WOString(20) & WOString(21)
ElseIf Z = 22 Then
WOString = WOString(1) & WOString(2) & WOString(3) & WOString(4) & WOString(5) & WOString(6) & _
WOString(7) & WOString(8) & WOString(9) & WOString(10) & WOString(11) & WOString(12) & WOString(13) & _
WOString(14) & WOString(15) & WOString(16) & WOString(17) & WOString(18) & WOString(19) & _
WOString(20) & WOString(21) & WOString(22)
ElseIf Z = 23 Then
WOString = WOString(1) & WOString(2) & WOString(3) & WOString(4) & WOString(5) & WOString(6) & _
WOString(7) & WOString(8) & WOString(9) & WOString(10) & WOString(11) & WOString(12) & WOString(13) & _
WOString(14) & WOString(15) & WOString(16) & WOString(17) & WOString(18) & WOString(19) & _
WOString(20) & WOString(21) & WOString(22) & WOString(23)
End If

'This section sends the email
Email_Subject = Range("C2") & " Shift " & Range("B2") & ": End of Shift Status" 'Email from respective sheet spvsr

'Email_Send_From = ""  Keep blank
Email_Send_To = Range("I2")
'Email_Cc = "Judy.Worthington@Navistar.com" *Remove REM when released
'Email_Bcc = "Jim.Borntrager@Navistar.com" *Remove REM when released

'Body strung together above is placed here
Email_Body = "FT Runners:" & Chr(13) & Chr(10) & FTString & Chr(13) & Chr(10) & _
  "Non-Field Test Runners:" & Chr(13) & Chr(10) & _
  NFTString & Chr(13) & Chr(10) & _
  "Work Orders:" & Chr(13) & Chr(10) & WOString

On Error GoTo debugs

Set Mail_Object = CreateObject("Outlook.Application")
Set Mail_Single = Mail_Object.CreateItem(0)
With Mail_Single
 .Subject = Email_Subject
 .To = Email_Send_To
 .CC = Email_Cc
 .BCC = Email_Bcc
 .Body = Email_Body
 .Send
End With

If Err.Description <> "" Then MsgBox Err.Description

debugs:

'If email recipient was incorrect
Else
  MsgBox "Select proper Email recipient (yourself) and try again."
    
End If

TheEnd:

End Sub


