Attribute VB_Name = "Hide_Unhide_Sheets"
Sub UnhideProtect()

' Unhide all schedule sheets but leave PROTECTED

    Sheet96.Select
    Sheet1.Visible = True
    Sheet1.Protect
    Sheet2.Visible = True
    Sheet2.Protect
    Sheet3.Visible = True
    Sheet3.Protect
    Sheet4.Visible = True
    Sheet4.Protect
    Sheet5.Visible = True
    Sheet5.Protect
    Sheet6.Visible = True
    Sheet6.Protect
    Sheet7.Visible = True
    Sheet7.Protect
    Sheet8.Visible = True
    Sheet8.Protect
    Sheet9.Visible = True
    Sheet9.Protect
    Sheet10.Visible = True
    Sheet10.Protect
    Sheet11.Visible = True
    Sheet11.Protect
    Sheet12.Visible = True
    Sheet12.Protect
    Sheet13.Visible = True
    Sheet13.Protect
    Sheet14.Visible = True
    Sheet14.Protect
    Sheet15.Visible = True
    Sheet15.Protect
    Sheet16.Visible = True
    Sheet16.Protect
    Sheet17.Visible = True
    Sheet17.Protect
    Sheet18.Visible = True
    Sheet18.Protect
    Sheet19.Visible = True
    Sheet19.Protect
    Sheet20.Visible = True
    Sheet20.Protect
    Sheet21.Visible = True
    Sheet21.Unprotect
    Sheet22.Visible = True
    Sheet22.Unprotect
    Sheet23.Visible = True
    Sheet23.Unprotect
    
End Sub

Sub UnhideUnprotect()

' Unhide AND Unprotect all schedule sheets

    Sheet96.Select
    Sheet1.Visible = True
    Sheet1.Unprotect
    Sheet2.Visible = True
    Sheet2.Unprotect
    Sheet3.Visible = True
    Sheet3.Unprotect
    Sheet4.Visible = True
    Sheet4.Unprotect
    Sheet5.Visible = True
    Sheet5.Unprotect
    Sheet6.Visible = True
    Sheet6.Unprotect
    Sheet7.Visible = True
    Sheet7.Unprotect
    Sheet8.Visible = True
    Sheet8.Unprotect
    Sheet9.Visible = True
    Sheet9.Unprotect
    Sheet10.Visible = True
    Sheet10.Unprotect
    Sheet11.Visible = True
    Sheet11.Unprotect
    Sheet12.Visible = True
    Sheet12.Unprotect
    Sheet13.Visible = True
    Sheet13.Unprotect
    Sheet14.Visible = True
    Sheet14.Unprotect
    Sheet15.Visible = True
    Sheet15.Unprotect
    Sheet16.Visible = True
    Sheet16.Unprotect
    Sheet17.Visible = True
    Sheet17.Unprotect
    Sheet18.Visible = True
    Sheet18.Unprotect
    Sheet19.Visible = True
    Sheet19.Unprotect
    Sheet20.Visible = True
    Sheet20.Unprotect
    Sheet21.Visible = True
    Sheet21.Unprotect
    Sheet22.Visible = True
    Sheet22.Unprotect
    Sheet23.Visible = True
    Sheet23.Unprotect
    
End Sub


Sub Hide()

' Hide AND Protect all schedule sheets

    Sheet96.Select
    Sheet1.Visible = False
    Sheet1.Protect
    Sheet2.Visible = False
    Sheet2.Protect
    Sheet3.Visible = False
    Sheet3.Protect
    Sheet4.Visible = False
    Sheet4.Protect
    Sheet5.Visible = False
    Sheet5.Protect
    Sheet6.Visible = False
    Sheet6.Protect
    Sheet7.Visible = False
    Sheet7.Protect
    Sheet8.Visible = False
    Sheet8.Protect
    Sheet9.Visible = False
    Sheet9.Protect
    Sheet10.Visible = False
    Sheet10.Protect
    Sheet11.Visible = False
    Sheet11.Protect
    Sheet12.Visible = False
    Sheet12.Protect
    Sheet13.Visible = False
    Sheet13.Protect
    Sheet14.Visible = False
    Sheet14.Protect
    Sheet15.Visible = False
    Sheet15.Protect
    Sheet16.Visible = False
    Sheet16.Protect
    Sheet17.Visible = False
    Sheet17.Protect
    Sheet18.Visible = False
    Sheet18.Protect
    Sheet19.Visible = False
    Sheet19.Protect
    Sheet20.Visible = False
    Sheet20.Protect
    Sheet21.Visible = False
    Sheet21.Protect
    Sheet22.Visible = False
    Sheet22.Protect
    Sheet23.Visible = False
    Sheet23.Protect

End Sub

Sub UnhideMT()

' Unhide Monday 2nd thru Tuesday 1st
    
    TodayDate = Range("TodayDate")
    SchedDate = Range("F3")
    
Call Hide

    Sheet1.Visible = True
    Sheet1.Protect
    Sheet2.Visible = True
    Sheet2.Protect
    Sheet3.Visible = True
    Sheet3.Protect
    Sheet4.Visible = True
    Sheet4.Protect
    Sheet5.Visible = True
    Sheet5.Protect
    
    Sheet96.Select
    If TodayDate > SchedDate Then End
    
    Sheet3.Select
    Call ResetTopSheet
    Range("ColumnNumber") = 8
    Call PreviewShift
    Sheet4.Select
    Call ResetTopSheet
    Range("ColumnNumber") = 9
    Call PreviewShift
    Sheet5.Select
    Call ResetTopSheet
    Range("ColumnNumber") = 10
    Call PreviewShift
           
End Sub

Sub UnhideTW()

' Unhide Tuesday 2nd thru Wednesday 1st

    TodayDate = Range("TodayDate")
    SchedDate = Range("I3")
    
    Call Hide
    
    Sheet5.Visible = True
    Sheet5.Protect
    Sheet6.Visible = True
    Sheet6.Protect
    Sheet7.Visible = True
    Sheet7.Protect
    Sheet8.Visible = True
    Sheet8.Protect
    
    Sheet96.Select
    If TodayDate > SchedDate Then End
    
    Sheet6.Select
    Call ResetTopSheet
    Range("ColumnNumber") = 11
    Call PreviewShift
    Sheet7.Select
    Call ResetTopSheet
    Range("ColumnNumber") = 12
    Call PreviewShift
    Sheet8.Select
    Call ResetTopSheet
    Range("ColumnNumber") = 13
    Call PreviewShift
 
End Sub

Sub UnhideWT()

' Unhide Wednesday 2nd thru Thursday 1st

    SchedDate = Range("L3")
    TodayDate = Range("TodayDate")

    Call Hide
    
    Sheet8.Visible = True
    Sheet8.Protect
    Sheet9.Visible = True
    Sheet9.Protect
    Sheet10.Visible = True
    Sheet10.Protect
    Sheet11.Visible = True
    Sheet11.Protect
    
    Sheet96.Select
    If TodayDate > SchedDate Then End
    
    Sheet9.Select
    Call ResetTopSheet
    Range("ColumnNumber") = 14
    Call PreviewShift
    Sheet10.Select
    Call ResetTopSheet
    Range("ColumnNumber") = 15
    Call PreviewShift
    Sheet11.Select
    Call ResetTopSheet
    Range("ColumnNumber") = 16
    Call PreviewShift
    
End Sub

Sub UnhideTF()

' Unhide Thursday 2nd thru Friday 1st

    TodayDate = Range("TodayDate")
    SchedDate = Range("O3")

    Call Hide
    
    Sheet11.Visible = True
    Sheet11.Protect
    Sheet12.Visible = True
    Sheet12.Protect
    Sheet13.Visible = True
    Sheet13.Protect
    Sheet14.Visible = True
    Sheet14.Protect
    
    Sheet96.Select
    If TodayDate > SchedDate Then End
    
    Sheet12.Select
    Call ResetTopSheet
    Range("ColumnNumber") = 17
    Call PreviewShift
    Sheet13.Select
    Call ResetTopSheet
    Range("ColumnNumber") = 18
    Call PreviewShift
    Sheet14.Select
    Call ResetTopSheet
    Range("ColumnNumber") = 19
    Call PreviewShift
 
End Sub
Sub UnhideFS()

' Unhide Friday 2nd thru Saturday 1st

    TodayDate = Range("TodayDate")
    SchedDate = Range("R3")

    Call Hide
    
    Sheet14.Visible = True
    Sheet14.Protect
    Sheet15.Visible = True
    Sheet15.Protect
    Sheet16.Visible = True
    Sheet16.Protect
    Sheet17.Visible = True
    Sheet17.Protect
    
    Sheet96.Select
    If TodayDate > SchedDate Then End
    
    Sheet15.Select
    Call ResetTopSheet
    Range("ColumnNumber") = 20
    Call PreviewShift
    Sheet16.Select
    Call ResetTopSheet
    Range("ColumnNumber") = 21
    Call PreviewShift
    Sheet17.Select
    Call ResetTopSheet
    Range("ColumnNumber") = 22
    Call PreviewShift

End Sub

Sub UnhideSS()

' Unhide Saturday 2nd thru Sunday 1st

    TodayDate = Range("TodayDate")
    SchedDate = Range("U3")

    Call Hide
    
    Sheet17.Visible = True
    Sheet17.Protect
    Sheet18.Visible = True
    Sheet18.Protect
    Sheet19.Visible = True
    Sheet19.Protect
    Sheet20.Visible = True
    Sheet20.Protect
    
    Sheet96.Select
    If TodayDate > SchedDate Then End
    
    Sheet18.Select
    Call ResetTopSheet
    Range("ColumnNumber") = 23
    Call PreviewShift
    Sheet19.Select
    Call ResetTopSheet
    Range("ColumnNumber") = 24
    Call PreviewShift
    Sheet20.Select
    Call ResetTopSheet
    Range("ColumnNumber") = 25
    Call PreviewShift
  
End Sub

Sub UnhideSM()

' Unhide Sunday 2nd thru (New) Monday 1st

    TodayDate = Range("TodayDate")
    SchedDate = Range("X3")

    Call Hide

    
    Sheet20.Visible = True
    Sheet20.Protect
    Sheet21.Visible = True
    Sheet21.Protect
    Sheet22.Visible = True
    Sheet22.Protect
    Sheet23.Visible = True
    Sheet23.Protect
    
    Sheet96.Select
    If TodayDate > SchedDate Then End
    
    Sheet21.Select
    Call ResetTopSheet
    Range("ColumnNumber") = 26
    Call PreviewShift
    Sheet22.Select
    Call ResetTopSheet
    Range("ColumnNumber") = 28
    Call PreviewShift
    Sheet23.Select
    Call ResetTopSheet
    Range("ColumnNumber") = 29
    Call PreviewShift
  
End Sub

Sub UnhideWeekend()

' Unhide Weekend Friday 2nd thru Monday 1st

' Unhide Friday 2nd thru Saturday 1st

    TodayDate = Range("TodayDate")
    SchedDate = Range("R3")

    Call Hide
    
    Sheet14.Visible = True
    Sheet14.Protect
    Sheet15.Visible = True
    Sheet15.Protect
    Sheet16.Visible = True
    Sheet16.Protect
    Sheet17.Visible = True
    Sheet17.Protect
    
    If TodayDate > SchedDate Then GoTo SAT
    
    Sheet15.Select
    Call ResetTopSheet
    Range("ColumnNumber") = 20
    Call PreviewShift
    Sheet16.Select
    Call ResetTopSheet
    Range("ColumnNumber") = 21
    Call PreviewShift
    Sheet17.Select
    Call ResetTopSheet
    Range("ColumnNumber") = 22
    Call PreviewShift
            
' Unhide Saturday 2nd thru Sunday 1st
SAT:
    TodayDate = Range("TodayDate")
    SchedDate = Range("U3")

    Call Hide
    
    Sheet17.Visible = True
    Sheet17.Protect
    Sheet18.Visible = True
    Sheet18.Protect
    Sheet19.Visible = True
    Sheet19.Protect
    Sheet20.Visible = True
    Sheet20.Protect
    
    If TodayDate > SchedDate Then GoTo SUN
    
    Sheet18.Select
    Call ResetTopSheet
    Range("ColumnNumber") = 23
    Call PreviewShift
    Sheet19.Select
    Call ResetTopSheet
    Range("ColumnNumber") = 24
    Call PreviewShift
    Sheet20.Select
    Call ResetTopSheet
    Range("ColumnNumber") = 25
    Call PreviewShift
            
' Unhide Sunday 2nd thru (New) Monday 1st
SUN:
    TodayDate = Range("TodayDate")
    SchedDate = Range("X3")

    Call Hide
    
    Sheet14.Visible = True
    Sheet14.Protect
    Sheet15.Visible = True
    Sheet15.Protect
    Sheet16.Visible = True
    Sheet16.Protect
    Sheet17.Visible = True
    Sheet17.Protect
    Sheet18.Visible = True
    Sheet18.Protect
    Sheet19.Visible = True
    Sheet19.Protect
    Sheet20.Visible = True
    Sheet20.Protect
    Sheet21.Visible = True
    Sheet21.Protect
    Sheet22.Visible = True
    Sheet22.Protect
    Sheet23.Visible = True
    Sheet23.Protect
    
    If TodayDate > SchedDate Then End
    
    Sheet21.Select
    Call ResetTopSheet
    Range("ColumnNumber") = 26
    Call PreviewShift
    Sheet22.Select
    Call ResetTopSheet
    Range("ColumnNumber") = 28
    Call PreviewShift
    Sheet23.Select
    Call ResetTopSheet
    Range("ColumnNumber") = 29
    Call PreviewShift
           
End Sub
