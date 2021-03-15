Attribute VB_Name = "modKeyState"
Option Explicit
Option Compare Text
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' modKeyState
' By Chip Pearson, www.cpearson.com, chip@cpearson.com
' This module contains functions for testing the state of the SHIFT, ALT, and CTRL
' keys.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


''''''''''''''''''''''''''''''''''''''''''''''''''''
' Declaration of GetKeyState API function. This
' tests the state of a specified key.
''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Declare Function GetKeyState Lib "user32" ( _
    ByVal nVirtKey As Long) As Integer
    
''''''''''''''''''''''''''''''''''''''''''
' This constant is used in a bit-wise AND
' operation with the result of GetKeyState
' to determine if the specified key is
' down.
''''''''''''''''''''''''''''''''''''''''''
Private Const KEY_MASK As Integer = &HFF80 ' decimal -128

'''''''''''''''''''''''''''''''''''''''''
' KEY CONSTANTS. Values taken
' from VC++ 6.0 WinUser.h file.
'''''''''''''''''''''''''''''''''''''''''
Private Const VK_LSHIFT = &HA0
Private Const VK_RSHIFT = &HA1
Private Const VK_LCONTROL = &HA2
Private Const VK_RCONTROL = &HA3
Private Const VK_LMENU = &HA4
Private Const VK_RMENU = &HA5
'''''''''''''''''''''''''''''''''''''''''
' The following four constants simply
' provide other names, CTRL and ALT,
' for CONTROL and MENU. "CTRL" and
' "ALT" are more familiar than
' "CONTROL" and "MENU". These constants
' provide no additional functionality.
' They simply provide more familiar
' names.
'''''''''''''''''''''''''''''''''''''''''
Private Const VK_LALT = VK_LMENU
Private Const VK_RALT = VK_RMENU
Private Const VK_LCTRL = VK_LCONTROL
Private Const VK_RCTRL = VK_RCONTROL

 ''''''''''''''''''''''''''''''''''''''''''''
' The following constants are used to specify,
' when testing CTRL, ALT, or SHIFT, whether
' the Left key, the Right key, either the
' Left OR Right, or BOTH the Left AND Right
' key is down.
'
' By default, the key-test procedures make
' no distinction between the Left and Right
' keys.
''''''''''''''''''''''''''''''''''''''''''''
Public Const BothLeftAndRightKeys = 0   ' Note: Bit-wise AND of LeftKey and RightKey
Public Const LeftKey = 1
Public Const RightKey = 2
Public Const LeftKeyOrRightKey = 3      ' Note: Bit-wise OR of LeftKey and RightKey


Public Function IsShiftKeyDown(Optional LeftOrRightKey As Long = LeftKeyOrRightKey)
''''''''''''''''''''''''''''''''''''''''''''''''
' IsShiftKeyDown
' Returns TRUE or FALSE indicating whether the
' SHIFT key is down.
'
' If LeftOrRightKey is omitted or LeftKeyOrRightKey,
' the function return TRUE if either the left or the
' right SHIFT key is down. If LeftKeyOrRightKey is
' LeftKey, then only the Left SHIFT key is tested.
' If LeftKeyOrRightKey is RightKey, only the Right
' SHIFT key is tested. If LeftOrRightKey is
' BothLeftAndRightKeys, the codes tests whether
' both the Left and Right keys are down. The default
' is to test for either Left or Right, making no
' distiction between Left and Right.
''''''''''''''''''''''''''''''''''''''''''''''''
    Dim Res As Long
    
    Select Case LeftOrRightKey
        Case LeftKey
            Res = GetKeyState(VK_LSHIFT) And KEY_MASK
        Case RightKey
            Res = GetKeyState(VK_RSHIFT) And KEY_MASK
        Case BothLeftAndRightKeys
            Res = (GetKeyState(VK_LSHIFT) And GetKeyState(VK_RSHIFT) And KEY_MASK)
        Case Else
            Res = GetKeyState(vbKeyShift) And KEY_MASK
    End Select
    
    IsShiftKeyDown = CBool(Res)
End Function

Public Function IsControlKeyDown(Optional LeftOrRightKey As Long = LeftKeyOrRightKey) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''
' IsControlKeyDown
' Returns TRUE or FALSE indicating whether the
' CTRL key is down.
'
' If LeftOrRightKey is omitted or LeftKeyOrRightKey,
' the function return TRUE if either the left or the
' right CTRL key is down. If LeftKeyOrRightKey is
' LeftKey, then only the Left CTRL key is tested.
' If LeftKeyOrRightKey is RightKey, only the Right
' CTRL key is tested. If LeftOrRightKey is
' BothLeftAndRightKeys, the codes tests whether
' both the Left and Right keys are down. The default
' is to test for either Left or Right, making no
' distiction between Left and Right.
''''''''''''''''''''''''''''''''''''''''''''''''
    Dim Res As Long
    
    Select Case LeftOrRightKey
        Case LeftKey
            Res = GetKeyState(VK_LCTRL) And KEY_MASK
        Case RightKey
            Res = GetKeyState(VK_RCTRL) And KEY_MASK
        Case BothLeftAndRightKeys
            Res = (GetKeyState(VK_LCTRL) And GetKeyState(VK_RCTRL) And KEY_MASK)
        Case Else
            Res = GetKeyState(vbKeyControl) And KEY_MASK
    End Select
    
    IsControlKeyDown = CBool(Res)

End Function

Public Function IsAltKeyDown(Optional LeftOrRightKey As Long = LeftKeyOrRightKey) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''
' IsAltKeyDown
' Returns TRUE or FALSE indicating whether the
' ALT key is down.
'
' If LeftOrRightKey is omitted or LeftKeyOrRightKey,
' the function return TRUE if either the left or the
' right ALT key is down. If LeftKeyOrRightKey is
' LeftKey, then only the Left ALT key is tested.
' If LeftKeyOrRightKey is RightKey, only the Right
' ALT key is tested. If LeftOrRightKey is
' BothLeftAndRightKeys, the codes tests whether
' both the Left and Right keys are down. The default
' is to test for either Left or Right, making no
' distiction between Left and Right.
''''''''''''''''''''''''''''''''''''''''''''''''
    Dim Res As Long
    
    Select Case LeftOrRightKey
        Case LeftKey
            Res = GetKeyState(VK_LALT) And KEY_MASK
        Case RightKey
            Res = GetKeyState(VK_RALT) And KEY_MASK
        Case BothLeftAndRightKeys
            Res = (GetKeyState(VK_LALT) And GetKeyState(VK_RALT) And KEY_MASK)
        Case Else
            Res = GetKeyState(vbKeyMenu) And KEY_MASK
    End Select
    
    IsAltKeyDown = CBool(Res)

End Function


Sub Test()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Test
' This is a procedrue to test and demonstrate the Key-State
' functions above. Since you can't run a macro in the VBA
' Editor if the SHIFT, ALT, or CTRL key is down, this procedure
' uses OnTime to execute the ProcTest test procedure. OnTime
' will call ProcTest two seconds after running this Test
' procedure. Immediately after executing Test, press the
' key(s) (Left/Right SHIFT, ALT, or CTRL) you want to test
' for. The procedure called by OnTime, ProcTest, displays the
' status of the Left/Right SHIFT, ALT, and CTRL keys.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Application.OnTime Now + TimeSerial(0, 0, 2), "ProcTest", , True
End Sub


Sub ProcTest()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ProcTest
' This procedure simply displays the status of the Left adn Right
' SHIFT, ALT, and CTRL keys.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Debug.Print "SHIFT KEY: ", "LEFT: " & CStr(IsShiftKeyDown(LeftKey)), _
                           "RIGHT: " & CStr(IsShiftKeyDown(RightKey)), _
                           "EITHER: " & CStr(IsShiftKeyDown(LeftKeyOrRightKey)), _
                           "BOTH:   " & CStr(IsShiftKeyDown(BothLeftAndRightKeys))
                        
Debug.Print "ALT KEY:   ", "LEFT: " & CStr(IsAltKeyDown(LeftKey)), _
                           "RIGHT: " & CStr(IsAltKeyDown(RightKey)), _
                           "EITHER: " & CStr(IsAltKeyDown(LeftKeyOrRightKey)), _
                           "BOTH:   " & CStr(IsAltKeyDown(BothLeftAndRightKeys))
                        
Debug.Print "CTRL KEY:   ", "LEFT: " & CStr(IsControlKeyDown(LeftKey)), _
                            "RIGHT: " & CStr(IsControlKeyDown(RightKey)), _
                            "EITHER: " & CStr(IsControlKeyDown(LeftKeyOrRightKey)), _
                            "BOTH:   " & CStr(IsControlKeyDown(BothLeftAndRightKeys))

End Sub

