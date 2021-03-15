<RequiredLevel:3>
<Boolean:IsForm> = DCount("*", "MSysObjects", "Name='" & <String:Name> & "' And [Type] = -32768")