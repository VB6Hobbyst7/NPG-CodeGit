Attribute VB_Name = "mTempDir"
Option Explicit

#If VBA7 And Win64 Then
  Declare PtrSafe Function GetTempPathA Lib "kernel32" _
  (ByVal lpBuffer As String, ByVal nSize As Long) As Long
#Else
  Declare Function GetTempPathA Lib "kernel32" _
  (ByVal lpBuffer As String, ByVal nSize As Long) As Long
#End If

Function TempDir() As String
'Returns the Temp folder
    Dim WinPath As String * 255
    WinPath = Space(255)
    TempDir = Left(WinPath, GetTempPathA(WinPath, Len(WinPath)))
End Function

