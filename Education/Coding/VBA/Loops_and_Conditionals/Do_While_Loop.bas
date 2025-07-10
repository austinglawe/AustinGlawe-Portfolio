' -----------------------------------------
' VBA Loops_Conditionals:
' Do While loop
' -----------------------------------------
'
' Purpose:
' - Repeat code while a condition is True.
' - Condition checked at the start of each iteration.
'
' Syntax:
'   Do While condition
'       ' Code
'   Loop
'
' Example 1:
'   Dim i As Integer
'   i = 1
'   Do While i <= 5
'       Debug.Print "Iteration " & i
'       i = i + 1
'   Loop
'
' Example 2:
'   Dim row As Long
'   row = 1
'   Do While Worksheets("Sheet1").Cells(row, 1).Value <> ""
'       Debug.Print Worksheets("Sheet1").Cells(row, 1).Value
'       row = row + 1
'   Loop
'
' Notes:
' - Ensure counter is incremented/decremented inside loop to prevent infinite loops.
' - "While ... Wend" exists but Do While ... Loop is preferred.
'
' Best practices:
' - Use for unknown number of iterations.
' - Condition checked before loop body runs; may run zero times.
'
' -----------------------------------------
