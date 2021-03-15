<RequiredLevel:3>
<Boolean:IsLinkedTable> = <Application.>DCount("*", "MSysObjects", "Name='" & <String:Name> & "' And [Type] = 6")