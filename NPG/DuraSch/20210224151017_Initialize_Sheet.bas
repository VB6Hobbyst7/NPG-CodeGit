Attribute VB_Name = "Initialize_Sheet"
Sub ResetTopSheet()
'This section clears runners, not Work Orders, and removes highlighted YELLOW cells from active sheet

ActiveSheet.Unprotect
    Range("A5:H12").Select
    Selection.ClearContents
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("A14:B31").Select
    Selection.ClearContents
    Range("E14:H31").Select
    Selection.ClearContents
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

    Range("H2").Select

ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
End Sub

Sub ResetSheet()
'This section clears all and removes highlighted YELLOW cells from active sheet

ActiveSheet.Unprotect
    Range("A5:H12").Select
    Selection.ClearContents
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("A14:B31").Select
    Selection.ClearContents
    Range("E14:H31").Select
    Selection.ClearContents
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    ActiveWindow.SmallScroll Down:=21
    Range("A33:I55").Select
    Selection.ClearContents
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("H33").Select
    Selection.FormulaR1C1 = " "
    Selection.AutoFill Destination:=Range("H33:H55"), Type:=xlFillDefault
    Range("H33:H55").Select
    ActiveWindow.SmallScroll Down:=-33
    
    Range("J5:J31").Select
    Selection.ClearContents

    Range("H2").Select

ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
End Sub

Sub ResetAllSheets()

    UnhideProtect

    Sheet1.Select
        Call ResetSheet
    Sheet2.Select
        Call ResetSheet
    Sheet3.Select
        Call ResetSheet
    Sheet4.Select
        Call ResetSheet
    Sheet5.Select
        Call ResetSheet
    Sheet6.Select
        Call ResetSheet
    Sheet7.Select
        Call ResetSheet
    Sheet8.Select
        Call ResetSheet
    Sheet9.Select
        Call ResetSheet
    Sheet10.Select
        Call ResetSheet
    Sheet11.Select
        Call ResetSheet
    Sheet12.Select
        Call ResetSheet
    Sheet13.Select
        Call ResetSheet
    Sheet14.Select
        Call ResetSheet
    Sheet15.Select
        Call ResetSheet
    Sheet16.Select
        Call ResetSheet
    Sheet17.Select
        Call ResetSheet
    Sheet18.Select
        Call ResetSheet
    Sheet19.Select
        Call ResetSheet
    Sheet20.Select
        Call ResetSheet
    Sheet21.Select
        Call ResetSheet
    Sheet22.Select
        Call ResetSheet
    Sheet23.Select
        Call ResetSheet
        
    Call UnhideMT
    
    Sheet96.Select

End Sub
Sub ResettingNewWeek()
'
' ResettingNewWeek Macro
'
    Call UnhideUnprotect

    Sheet1.Select
        Call ResetSheet
    Sheet2.Select
        Call ResetSheet

    Sheet1.Visible = True
    Sheet1.Unprotect
    Sheet2.Visible = True
    Sheet2.Unprotect

    Sheet22.Select
    Range("A5:J55").Copy
    
    Sheet1.Select
    Range("A5").Select
    ActiveSheet.Paste
    Range("k5:k31").Copy
    Range("k5").Select
    Selection.PasteSpecial Paste:=xlPasteValues
    Range("H2").Select
    
    Sheet23.Select
    Range("A5:J55").Copy
    Sheet2.Select
    Range("A5").Select
    ActiveSheet.Paste
    Range("K5:K31").Copy
    Range("K5").Select
    Selection.PasteSpecial Paste:=xlPasteValues
    Range("H2").Select
        
    Sheet3.Select
        Call ResetSheet
    Sheet4.Select
        Call ResetSheet
    Sheet5.Select
        Call ResetSheet
    Sheet6.Select
        Call ResetSheet
    Sheet7.Select
        Call ResetSheet
    Sheet8.Select
        Call ResetSheet
    Sheet9.Select
        Call ResetSheet
    Sheet10.Select
        Call ResetSheet
    Sheet11.Select
        Call ResetSheet
    Sheet12.Select
        Call ResetSheet
    Sheet13.Select
        Call ResetSheet
    Sheet14.Select
        Call ResetSheet
    Sheet15.Select
        Call ResetSheet
    Sheet16.Select
        Call ResetSheet
    Sheet17.Select
        Call ResetSheet
    Sheet18.Select
        Call ResetSheet
    Sheet19.Select
        Call ResetSheet
    Sheet20.Select
        Call ResetSheet
    Sheet21.Select
        Call ResetSheet
    Sheet22.Select
        Call ResetSheet
    Sheet23.Select
        Call ResetSheet
        
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
    Range("AB7:AC45").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("F7").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Range("H7:Z45").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    
    Range("AB7:AC45").Select
    Selection.ClearContents
    Range("F3").Select
    Selection.ClearContents
    
End Sub

