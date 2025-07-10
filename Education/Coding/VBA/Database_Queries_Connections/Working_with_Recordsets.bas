' -----------------------------------------
' VBA Database_Queries_Connections:
' Working with Recordsets
' -----------------------------------------
'
' Navigation:
' - .MoveFirst, .MoveLast, .MoveNext, .MovePrevious
' - .EOF, .BOF to detect bounds
'
' Example loop:
'   Do While Not rs.EOF
'       Debug.Print rs.Fields("FieldName").Value
'       rs.MoveNext
'   Loop
'
' Filtering:
'   rs.Filter = "Country = 'USA'"
'   rs.Filter = adFilterNone  ' Clear filter
'
' Editing records:
'   rs.MoveFirst
'   rs.Edit
'   rs.Fields("City").Value = "New York"
'   rs.Update
'
' Adding records:
'   rs.AddNew
'   rs.Fields("CustomerName").Value = "New Customer"
'   rs.Update
'
' Deleting records:
'   rs.Delete
'   rs.MoveNext
'
' Notes:
' - Cursor and lock types (e.g., adOpenKeyset, adLockOptimistic) affect editing.
' - Always check .EOF and .BOF.
' - Close Recordsets when done.
'
' -----------------------------------------
