'http://allenbrowne.com/ser-28.html
If IsNull(Me.cboShowCat) Then 'TODO 1. Set correct name of combobox, 2. Set correct Filter 
	Me.FilterOn = False
Else
	Me.Filter = "ProductCatID = """ & Me.cboShowCat & """"
	Me.FilterOn = True
End If