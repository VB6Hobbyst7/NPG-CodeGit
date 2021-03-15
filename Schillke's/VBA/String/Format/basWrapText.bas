Attribute VB_Name = "basWrapText"
'? WrapText ("The quick brown fox jumps over the lazy dog.", 10)
' WrapText
'http://www.utteraccess.com/wiki/index.php/WrapText RequiredLevel:3
' Code courtesy of UtterAccess Wiki
' Licensed under Creative Commons License
' http://creativecommons.org/licenses/by-sa/3.0/
'
' You are free to use this code in any application,
' provided this notice is left unchanged.
'
' rev  date                          brief descripton
' 1.0  2012-08-26                    wraps a given string of text to a certain number of characters (similar to a textbox)
'
Public Function WrapText(StringOfText As String, MaxLength As Long) As String
'wraps a given string of text to a certain number of characters (similar to a textbox)
'http://www.utteraccess.com/wiki/index.php/WrapText
   Dim MyArray() As String
   Dim Counter As Long
   Dim i As Long
   Dim j As Long
   Dim TempString As String

   StringOfText = Replace(StringOfText, vbCrLf, " ")
       
   MyArray = Split(StringOfText, " ")
   
   For i = 0 To UBound(MyArray)
       If Counter + Len(MyArray(i)) <= MaxLength Then      'line + new word is less than MaxLength
           WrapText = WrapText & " " & MyArray(i)          'add word to line
           Counter = Counter + Len(MyArray(i)) + 1
       ElseIf Len(MyArray(i)) > MaxLength Then             'single word longer than MaxLength
           For j = 1 To Len(MyArray(i)) Step MaxLength     'split word as many times as necessary
               WrapText = WrapText & vbCrLf & Mid(MyArray(i), j, MaxLength)
           Next j
           Counter = Len(MyArray(i)) - j + 1 + MaxLength
       Else
           WrapText = WrapText & vbCrLf & MyArray(i)       'line + new word is more than MaxLength
           Counter = Len(MyArray(i)) + 1
       End If
   Next i
   
   WrapText = Trim(WrapText)
   If Left(WrapText, 2) = vbCrLf Then WrapText = Mid(WrapText, 3)
   
End Function
