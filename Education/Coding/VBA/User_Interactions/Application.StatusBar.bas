' -----------------------------------------
' Application.StatusBar overview and guidance
' -----------------------------------------
'
' Application.StatusBar allows you to write text to the lower-left status bar in Excel.
' This is useful for providing passive, non-intrusive feedback to users during code execution.
'
' Basic usage:
'   Application.StatusBar = "Processing your request..."
'
' Clearing it:
'   Application.StatusBar = False
'   ' This restores Excel's normal status messages.
'
' Behavior:
' - StatusBar is non-modal: your macro keeps running while text is displayed.
' - The message stays visible until explicitly cleared or Excel is restarted.
' - You can update it anytime by assigning a new string.
'
' Best practices:
' - Always clear the StatusBar at the end of your macro to avoid leaving stale messages.
'
' Example 1 (simple message):
'     Application.StatusBar = "Exporting data..."
'     ' [Your code runs]
'     Application.StatusBar = False
'
' Example 2 (handling errors and always cleaning up):
'     Sub Example_StatusBar()
'         On Error GoTo Cleanup
'         Application.StatusBar = "Running... Please wait."
'
'         ' Simulate work:
'         Dim i As Long
'         For i = 1 To 10000000
'             ' Simulated work loop
'         Next i
'
' Cleanup:
'         Application.StatusBar = False
'     End Sub
'
' Usage scenarios:
' - Inform the user about long-running processes ("Loading report...", "Saving file...").
' - Provide unobtrusive progress feedback.
'
' Limitations:
' - No user interaction (display-only).
' - No formatting (text only).
' - Forgetting to clear it will leave your message stuck in the status bar.
'
' Notes:
' - Works well alongside Application.ScreenUpdating = False for smooth UX.
'
' -----------------------------------------
