' -----------------------------------------
' VBA Error_Handling overview
' -----------------------------------------
'
' Purpose:
' - Gracefully catch and handle run-time errors.
' - Provide custom responses or fail-safes instead of crashing.
'
' Main tools:
' 1. On Error Resume Next
' 2. On Error GoTo [label]
' 3. Err object (Err.Number, Err.Description)
' 4. Err.Clear
'
' Best practices:
' - Handle only expected errors deliberately.
' - Avoid leaving On Error Resume Next active without properly clearing it later.
' - Always reset error handler when done (On Error GoTo 0).
'
' -----------------------------------------

