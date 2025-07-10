' -----------------------------------------
' VBA Error_Handling:
' On Error Resume Next
' -----------------------------------------
'
' Purpose:
' - Ignore runtime errors and continue execution with next line.
'
' Syntax:
'   On Error Resume Next
'
' Example:
'   On Error Resume Next
'   Kill "C:\Test\nonexistentfile.txt"  ' May error if file missing
'   On Error GoTo 0  ' Restore default error handling
'
' Important:
' - Use sparingly to avoid hiding bugs.
' - Always restore error handling with On Error GoTo 0 after use.
'
' Example with object assignment:
'   On Error Resume Next
'   Dim ws As Worksheet
'   Set ws = Worksheets("SheetDoesNotExist")
'   If ws Is Nothing Then MsgBox "Worksheet not found"
'   On Error GoTo 0
'
' Best practices:
' - Use only for small blocks where errors are expected.
' - Check results explicitly after error-prone calls.
' - Never leave it active for large sections of code.
'
' -----------------------------------------
