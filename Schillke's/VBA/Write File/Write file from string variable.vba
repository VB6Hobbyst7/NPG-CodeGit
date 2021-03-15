Dim iFile As Integer: iFile = FreeFile
Open <String:Filename> For Output Access Write As #iFile
Print #iFile, <String:Output>    
Close #iFile