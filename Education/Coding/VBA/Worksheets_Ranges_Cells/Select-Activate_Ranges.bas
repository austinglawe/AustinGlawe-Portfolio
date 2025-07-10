' -----------------------------------------
' VBA Worksheets_Ranges_Cells:
' Selecting and activating ranges
' -----------------------------------------
'
' Purpose:
' - Select or activate objects for user-visible interaction.
'
' Syntax:
' - Worksheet activation:
'     Worksheets("Sheet1").Activate
'
' - Range selection:
'     Worksheets("Sheet1").Range("A1:B2").Select
'
' Notes:
' - Selection.Value allows referring to currently selected range.
'
' Example (inefficient style):
'   Worksheets("Sheet1").Activate
'   Range("A1").Select
'   Selection.Value = "Hello"
'
' Example (preferred style):
'   Worksheets("Sheet1").Range("A1").Value = "Hello"
'
' Best practices:
' - Avoid Select and Activate unless absolutely necessary for UI presentation.
' - Direct object references are faster, safer, and easier to maintain.
' - Only use Select if visually highlighting cells for the user is required.
'
' -----------------------------------------
