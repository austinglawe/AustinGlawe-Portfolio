' -----------------------------------------
' VBA Loops_Conditionals:
' Do Until loop
' -----------------------------------------
'
' Purpose:
' - Repeat code until condition is True.
' - Condition checked before loop starts.
'
' Syntax:
'   Do Until condition
'       ' Code
'   Loop
'
' Example 1:
'   Dim i As Integer
'   i = 1
'   Do Until i > 5
'       Debug.Print "Iteration " & i
'       i = i + 1
'   Loop
'
' Example 2:
'   Dim row As Long
'   row = 1
'   Do Until Worksheets("Sheet1").Cells(row, 1).Value = ""
'       Debug.Print Worksheets("Sheet1").Cells(row, 1).Value
'       row = row + 1
'   Loop
'
' Alternative:
'   Do
'       ' Code
'   Loop Until condition
'   ' Condition checked at end — loop runs at least once.
'
' Best practices:
' - Do Until is equivalent to Do While Not.
' - Ensure exit condition to prevent infinite loops.
' - Ideal for cases where loop stops once a condition is True.
'
' -----------------------------------------
