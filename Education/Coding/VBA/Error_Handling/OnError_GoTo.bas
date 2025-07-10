' -----------------------------------------
' VBA Error_Handling:
' On Error GoTo [label]
' -----------------------------------------
'
' Purpose:
' - Jump to a specific error handling block when a runtime error occurs.
'
' Syntax:
'   On Error GoTo ErrHandler
'   ' Code
'   Exit Sub
' ErrHandler:
'   ' Error handling code
'   MsgBox "Error #" & Err.Number & ": " & Err.Description
'   Resume Next  ' or Resume
'
' Example:
'   Sub Example()
'       On Error GoTo ErrHandler
'       Dim x As Integer
'       x = 1 / 0
'       Exit Sub
'   ErrHandler:
'       MsgBox "Error #" & Err.Number & ": " & Err.Description
'       Resume Next
'   End Sub
'
' Best practices:
' - Always put Exit Sub before the error handler.
' - Use error handler to log, clean up, or notify.
' - Use Resume or Resume Next to control flow after handling.
'
' -----------------------------------------
