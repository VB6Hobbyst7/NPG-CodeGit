For Each <Worksheet> In <Application.>ThisWorkbook.Worksheets
	<Worksheet>.Protect UserInterfaceOnly:=True
	'<Worksheet>.Protect , , , , True
Next ws