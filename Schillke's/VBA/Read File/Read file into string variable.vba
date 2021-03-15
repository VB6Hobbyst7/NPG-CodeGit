Dim iFile As Integer: iFile = FreeFile
Open <String:Filename> For Input As #iFile
<String:FileContent> = Input(LOF(iFile), iFile)       
Close #iFile