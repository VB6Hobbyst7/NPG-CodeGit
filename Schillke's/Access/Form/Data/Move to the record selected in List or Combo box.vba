Me.RecordsetClone.FindFirst "[<cursor>] = " & Me![] 'TODO: 1. set field for FindFirst and 2. set Combo or Listbox name in Me![]
Me.Bookmark = Me.RecordsetClone.Bookmark