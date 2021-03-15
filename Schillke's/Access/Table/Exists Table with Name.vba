<RequiredLevel:3>
<Boolean:IsTable> = <Application.>DCount("*", "MSysObjects", "Name='" & <String:Name> & "' And [Type] = 1")