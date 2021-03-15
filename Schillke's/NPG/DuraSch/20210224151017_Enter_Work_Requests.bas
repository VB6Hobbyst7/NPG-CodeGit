Attribute VB_Name = "Enter_Work_Requests"
Sub EnterWO()             ' Enter Work Requests
Attribute EnterWO.VB_ProcData.VB_Invoke_Func = " \n14"

    Priority = Range("Priority")
    ProjVeh = Range("ProjVeh")
    ChargeNumber = Range("ChargeNumber")
    ProjectDesc = Range("ProjectDesc")
    'EstHours = Range("EstHours")
    WONumber = Range("WONumber")
    
    If Priority = "" Then End         'End if any blanks in WO copy location
    If ProjVeh = "" Then End
    If ChargeNumber = "" Then End
    If ProjectDesc = "" Then End
    'If EstHours = "" Then End
    If WONumber = "" Then End
    
    For a = 18 To 38                  'x = Row number to define sheets
        
        If Cells(a, 15).Value = True Then
            Else: GoTo Increment
        End If
    
        SheetNumber = Cells(a, 14).Value
        Sheets(SheetNumber).Activate  'Move to selected sheet to paste WO information
        
        For x = 33 To 55
            Cells(x, 1).Select
            If Cells(x, 1) = "" Then GoTo ENTER
        Next x
    
ENTER:                                'Enters data into selected sheet
        Cells(x, 1).Value = Priority
        Cells(x, 2).Value = ProjVeh
        Cells(x, 3).Value = ChargeNumber
        Cells(x, 4).Value = ProjectDesc
        'Cells(x, 8).Value = EstHours
        Cells(x, 9).Value = WONumber
        

        
        Sheets("Enter Work Orders").Select
    
Increment:
    Next a

    ActiveSheet.Unprotect
    Range("A18:E18").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection.Font
        .Name = "Arial"
        .FontStyle = "Regular"
        .Size = 8
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
Selection.Locked = False
    Selection.FormulaHidden = False
    
    Range("A18").Select
    
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    

    
End Sub
Sub EnterHouseWO()             ' Enter House Work Requests

    On Error GoTo TheEnd

RowNumber = ActiveCell.Row

Dim answer As Integer
answer = MsgBox("Are you sure you want to enter House Work Order  " & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Cells(RowNumber, 2).Value & " ?", vbYesNo + vbQuestion, "CONFIRM HOUSE WORK REQUEST")
If answer = vbYes Then
     GoTo P1:
Else
     GoTo TheEnd:
End If

P1:
    HPriority = Cells(RowNumber, 1).Value
    HProjVeh = Cells(RowNumber, 2).Value
    HChargeNumber = Cells(RowNumber, 3).Value
    HProjectDesc = Cells(RowNumber, 4).Value
    HWONumber = Cells(RowNumber, 9).Value
    
    For a = 13 To 33                  'a = Column number to define sheets
        
        If Cells(3, a).Value = True Then
            Else: GoTo Increment
        End If
    
        SheetNumber = Cells(2, a).Value
        Sheets(SheetNumber).Activate  'Move to selected sheet to paste WO information
        
        For x = 33 To 55
            Cells(x, 2).Select
            If Cells(x, 2) = "" Then GoTo ENTER
        Next x
    
ENTER:                                'Enters data into selected sheet
        Cells(x, 1).Value = HPriority
        Cells(x, 2).Value = HProjVeh
        Cells(x, 3).Value = HChargeNumber
        Cells(x, 4).Value = HProjectDesc
        Cells(x, 9).Value = HWONumber
        
        Sheets("House Work Requests").Select
    
Increment:
    Next a

TheEnd:

End Sub


