Attribute VB_Name = "modAccessObject"
Public Function HasProperty(obj As Object, strPropName As String) As Boolean
'Test the object has the named property. RequiredLevel:4
    'Purpose: Return true if the object has the property.
    Dim varDummy As Variant

    On Error Resume Next
    varDummy = obj.Properties(strPropName)
    HasProperty = (Err.Number = 0)
End Function
