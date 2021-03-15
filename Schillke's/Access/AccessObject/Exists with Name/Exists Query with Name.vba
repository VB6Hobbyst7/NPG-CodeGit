<RequiredLevel:3>
<Boolean:IsQuery> = <Application.>DCount("*", "MSysObjects", "Name='" & <String:Name> & "' And [Type] = 5")