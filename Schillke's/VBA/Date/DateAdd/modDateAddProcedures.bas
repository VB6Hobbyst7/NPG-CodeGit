Attribute VB_Name = "modDateAddProcedures"
Option Explicit

'==========================================================
'https://support.microsoft.com/en-us/kb/115489 / adapted
' The DateAddWorkday() function provides a workday substitute
' for DateAdd("w", number, date). This function performs
' error checking and ignores fractional Interval values.
'==========================================================
Function DateAddWorkday(ByVal Dat As Date, ByVal Interval As Integer) As Date
'Add just the workdays (Monday through Friday), rather than weekdays (Sunday through Saturday), to a date.
'https://support.microsoft.com/kb/115489 RequiredLevel:3
    Dim Weeks As Long, OddDays As Long, Temp As String
    
    If VarType(Dat) <> 7 Or VarType(Interval) < 2 Or _
        VarType(Interval) > 5 Then
        DateAddWorkday = Dat
        ElseIf Interval = 0 Then
        DateAddWorkday = Dat
        ElseIf Interval > 0 Then
        Interval = Int(Interval)
        
        ' Make sure Dat is a workday (round down).
        Temp = Format(Dat, "ddd")
        If Temp = "Sun" Then
            Dat = Dat - 2
            ElseIf Temp = "Sat" Then
            Dat = Dat - 1
        End If
        
        ' Calculate Weeks and OddDays.
        Weeks = Int(Interval / 5)
        OddDays = Interval - (Weeks * 5)
        Dat = Dat + (Weeks * 7)
        
        ' Take OddDays weekend into account.
        If (DatePart("w", Dat) + OddDays) > 6 Then
            Dat = Dat + OddDays + 2
        Else
            Dat = Dat + OddDays
        End If
        
        DateAddWorkday = Dat
        Else                         ' Interval is < 0
            Interval = Int(-Interval) ' Make positive & subtract later.
            
            ' Make sure Dat is a workday (round up).
            Temp = Format(Dat, "ddd")
            If Temp = "Sun" Then
                Dat = Dat + 1
                ElseIf Temp = "Sat" Then
                Dat = Dat + 2
            End If
            
            ' Calculate Weeks and OddDays.
            Weeks = Int(Interval / 5)
            OddDays = Interval - (Weeks * 5)
            Dat = Dat - (Weeks * 7)
            
            ' Take OddDays weekend into account.
            If (DatePart("w", Dat) - OddDays) < 2 Then
                Dat = Dat - OddDays - 2
            Else
                Dat = Dat - OddDays
            End If
            
            DateAddWorkday = Dat
        End If
        
End Function

