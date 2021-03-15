Dim strInput As String
Dim booValidatedOk As Boolean: booValidatedOk = False
Do
    strInput = InputBox(Prompt:="Give Input")
    If Len(strInput) = 0 Then GoTo ExitValidateInput
    GoSub ValidateInput
    If Not booValidatedOk Then MsgBox "Input incorrect", vbExclamation
Loop While Not booValidatedOk
    
GoTo SkipValidateInput
ValidateInput:
    booValidatedOk = IsNumeric(strInput) 'TODO: replace by your validation
    Return
ExitValidateInput:
SkipValidateInput: