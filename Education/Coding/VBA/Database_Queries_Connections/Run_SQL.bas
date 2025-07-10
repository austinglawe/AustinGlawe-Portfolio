' -----------------------------------------
' VBA Database_Queries_Connections:
' Running SQL queries and retrieving results
' -----------------------------------------
'
' Purpose:
' - Execute SQL queries via ADO and read results using Recordsets.
'
' Example (late binding):
'   Dim conn As Object, rs As Object
'   Set conn = CreateObject("ADODB.Connection")
'   Set rs = CreateObject("ADODB.Recordset")
'
'   conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Path\Database.accdb;"
'   Dim sql As String
'   sql = "SELECT * FROM Customers WHERE Country='USA'"
'
'   rs.Open sql, conn, 1, 3  ' 1=adOpenKeyset, 3=adLockOptimistic
'
'   Do While Not rs.EOF
'       Debug.Print rs.Fields("CustomerName").Value
'       rs.MoveNext
'   Loop
'
'   ' Copy recordset to worksheet:
'   Worksheets("Sheet1").Range("A1").CopyFromRecordset rs
'
'   rs.Close
'   conn.Close
'   Set rs = Nothing
'   Set conn = Nothing
'
' Notes:
' - Close objects to free resources.
' - Use appropriate cursor and lock types.
' - Access fields by name.
' - Use CopyFromRecordset for fast data import.
'
' -----------------------------------------
