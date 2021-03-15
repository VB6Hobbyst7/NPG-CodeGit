<RequiredLevel:3><HelpUrl:http://codevba.com/office/delete_file_kill.htm>
If Len(Dir$(<String:FileToDelete>)) > 0 Then 
    SetAttr <String:FileToDelete>, vbNormal
    Kill <String:FileToDelete>
End If