Dim doc As Document
For Each doc In Documents
	doc.Close SaveChanges:=wdSaveChanges
Next