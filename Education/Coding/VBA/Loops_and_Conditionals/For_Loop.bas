' -----------------------------------------
' VBA Loops_Conditionals:
' For loop
' -----------------------------------------
'
' Purpose:
' - Repeat code a known number of times.
'
' Syntax:
'   For counter = start To end [Step increment]
'       ' Code to repeat
'   Next counter
'
' Examples:
'
' 1. Count up from 1 to 5:
'   Dim i As Integer
'   For i = 1 To 5
'       Debug.Print "Iteration " & i
'   Next i
'
' 2. Count down from 10 to 1:
'   For i = 10 To 1 Step -1
'       Debug.Print i
'   Next i
'
' 3. Custom step size:
'   For i = 2 To 10 Step 2
'       Debug.Print i  ' 2, 4, 6, 8, 10
'   Next i
'
' Best practices:
' - Use clear counter names for readability.
' - Prefer dynamic limits (e.g., lastRow) over hardcoded numbers.
' - Avoid unnecessary nested For loops.
'
' -----------------------------------------
