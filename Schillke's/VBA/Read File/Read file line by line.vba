Dim iFile As Integer: iFile = FreeFile
Open <String:Filename> For Input As #iFile
Do Until EOF(iFile)
    Line Input #iFile, <String:TextLine>
    <cursor>
Loop
Close #iFile