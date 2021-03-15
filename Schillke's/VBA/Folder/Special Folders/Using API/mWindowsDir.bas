Attribute VB_Name = "mWindowsDir"
Option Explicit

#If VBA7 And Win64 Then
  Declare PtrSafe Function GetWindowsDirectoryA Lib "kernel32" _
  (ByVal lpBuffer As String, ByVal nSize As Long) As Long
#Else
  Declare Function GetWindowsDirectoryA Lib "kernel32" _
  (ByVal lpBuffer As String, ByVal nSize As Long) As Long
#End If

Function WindowsDir() As String
'Returns the Windows folder
    Dim WinPath As String * 255
    WinPath = Space(255)
    WindowsDir = Left(WinPath, GetWindowsDirectoryA(WinPath, Len(WinPath)))
End Function

