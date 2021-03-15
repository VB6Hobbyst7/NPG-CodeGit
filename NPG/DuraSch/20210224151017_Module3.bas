Attribute VB_Name = "Module3"

Sub Macro5()
Attribute Macro5.VB_ProcData.VB_Invoke_Func = "m\n14"
'
' Macro5 Macro
'
' Keyboard Shortcut: Ctrl+m
'
    Range("I5").Select
    Selection.Copy
    Range("I5:I31").Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Range("I13").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Range("I5").Select
End Sub
Sub Macro6()
Attribute Macro6.VB_ProcData.VB_Invoke_Func = "k\n14"
'
' Macro6 Macro
'
' Keyboard Shortcut: Ctrl+k
'
    Range("I5:I31").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=""Needs Trained"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Range("I5:I31").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=""Signed Off"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 5296274
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Range("I5").Select
End Sub
