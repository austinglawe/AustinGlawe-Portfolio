' -----------------------------------------
' VBA FileSystem_Operations:
' Error handling and permission management
' -----------------------------------------
'
' Importance:
' - Handle errors from missing files, permission denied, locks.
'
' Example pattern:
'   On Error GoTo ErrHandler
'   ' File operations here
'   Exit Sub
' ErrHandler:
'   MsgBox "File system error #" & Err.Number & ": " & Err.Description
'   Resume Next
'
' Handling permissions:
' - Catch error 70 (Permission denied).
' - Check existence before operations.
' - Avoid system or protected files/folders.
'
' Tips:
' - Use Dir to check file/folder existence.
' - Turn off ScreenUpdating during bulk operations.
' - Run Excel as admin if necessary.
'
' Best practices:
' - Always use error handling.
' - Provide clear user messages.
' - Log errors for troubleshooting.
'
' -----------------------------------------
