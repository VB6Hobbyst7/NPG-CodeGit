Attribute VB_Name = "mod64bit"
Public Function Is64Bit() As Boolean
'Determines if Office version is 64 bit
'http://www.cpearson.com/excel/Bitness.aspx
#If VBA7 And Win64 Then
    Is64Bit = True
#Else
    Is64Bit = False
#End If
End Function
