Attribute VB_Name = "Module4"
Sub RefreshALL()
Attribute RefreshALL.VB_ProcData.VB_Invoke_Func = " \n14"
'
' RefreshALL Macro
'
ActiveWorkbook.RefreshALL
End Sub
Sub Macro7()
Attribute Macro7.VB_ProcData.VB_Invoke_Func = "i\n14"
'
' Macro7 Macro
'
' Keyboard Shortcut: Ctrl+i
'
    Range("L13").Select
    ActiveSheet.Buttons.Add(929.25, 167.25, 120, 27).Select
    ActiveSheet.Paste
    ActiveSheet.Shapes.Range(Array("Button 12")).Select
End Sub
