' -----------------------------------------
' VBA Database_Queries_Connections:
' Setting up ADO connections
' -----------------------------------------
'
' Purpose:
' - Connect to databases using ADO from VBA.
'
' Late binding example:
'   Dim conn As Object
'   Set conn = CreateObject("ADODB.Connection")
'   conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Path\Database.accdb;"
'
' Early binding example:
'   Dim conn As ADODB.Connection
'   Set conn = New ADODB.Connection
'   conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Path\Database.accdb;"
'   conn.Open
'
' Important:
' - Close connections after use:
'     conn.Close
'     Set conn = Nothing
'
' Best practices:
' - Use late binding for portability.
' - Ensure connection strings match database type/version.
' - Add error handling around connection code.
'
' -----------------------------------------
