' -----------------------------------------
' VBA Database_Queries_Connections:
' Connection strings basics
' -----------------------------------------
'
' Purpose:
' - Define parameters for connecting to databases.
'
' Key elements:
' - Provider: Database driver.
' - Data Source: Path or server.
' - User ID/Password: Authentication.
' - Additional options.
'
' Common examples:
' Access 2007+ (.accdb):
'   Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Path\Database.accdb;
'
' Access 2003 (.mdb):
'   Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Path\Database.mdb;
'
' SQL Server (Windows Auth):
'   Provider=SQLOLEDB;Data Source=SERVERNAME;Initial Catalog=DatabaseName;Integrated Security=SSPI;
'
' SQL Server (SQL Auth):
'   Provider=SQLOLEDB;Data Source=SERVERNAME;Initial Catalog=DatabaseName;User ID=userid;Password=pwd;
'
' Tips:
' - Use trusted connection when possible.
' - Store connection strings securely.
' - Test strings before use.
'
' -----------------------------------------
