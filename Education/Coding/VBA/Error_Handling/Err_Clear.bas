' -----------------------------------------
' VBA Error_Handling:
' Err.Clear
' -----------------------------------------
'
' Purpose:
' - Reset the Err object to clear error info.
'
' When to use:
' - After handling an error before continuing.
' - To avoid confusion between past and new errors.
'
' Example:
'   ErrHandler:
'       MsgBox "Error #" & Err.Number & ": " & Err.Description
'       Err.Clear
'       Resume Next
'
' Best practices:
' - Clear errors only in error handling blocks.
' - Avoid clearing errors outside handlers to prevent hiding errors.
'
' -----------------------------------------
