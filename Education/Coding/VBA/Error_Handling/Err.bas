' -----------------------------------------
' VBA Error_Handling:
' The Err object
' -----------------------------------------
'
' Purpose:
' - Holds details about the last runtime error.
'
' Key properties:
' - Err.Number: Numeric error code.
' - Err.Description: Text description of error.
' - Err.Source: Name of source of error.
' - Err.HelpFile: Path to help file.
' - Err.HelpContext: Help topic ID.
'
' Example usage:
'   ErrHandler:
'       MsgBox "Error #" & Err.Number & ": " & Err.Description & vbCrLf & _
'              "Source: " & Err.Source
'       Err.Clear
'       Resume Next
'
' Common error codes:
' - 11: Division by zero
' - 9: Subscript out of range
' - 13: Type mismatch
' - 1004: Application-defined or object-defined error
'
' Best practices:
' - Check Err.Number to determine error type.
' - Use Err.Description for meaningful messages.
' - Call Err.Clear after handling error.
' - Don’t silently ignore errors.
'
' -----------------------------------------
