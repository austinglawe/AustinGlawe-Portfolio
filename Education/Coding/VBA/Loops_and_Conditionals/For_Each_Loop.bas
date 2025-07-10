' -----------------------------------------
' VBA Loops_Conditionals:
' For Each loop
' -----------------------------------------
'
' Purpose:
' - Iterate over collections or object groups cleanly.
'
' Syntax:
'   For Each element In collection
'       ' Code
'   Next element
'
' Examples:
'
' 1. Loop worksheets:
'   Dim ws As Worksheet
'   For Each ws In ThisWorkbook.Worksheets
'       Debug.Print ws.Name
'   Next ws
'
' 2. Loop cells in a range:
'   Dim c As Range
'   For Each c In Worksheets("Sheet1").Range("A1:A5")
'       c.Value = c.Row
'   Next c
'
' 3. Loop files in folder (FSO):
'   Dim fso As Object, fld As Object, fil As Object
'   Set fso = CreateObject("Scripting.FileSystemObject")
'   Set fld = fso.GetFolder("C:\Test")
'
'   For Each fil In fld.Files
'       Debug.Print fil.Name
'   Next fil
'
' Best practices:
' - Prefer For Each for collections (Worksheets, Cells, Files, etc.).
' - Use appropriate variable types for clarity and performance.
' - Cleaner and safer than For/Next + index when working with collections.
'
' -----------------------------------------
