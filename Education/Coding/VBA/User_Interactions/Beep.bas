' -----------------------------------------
' Beep overview and guidance
' -----------------------------------------
'
' The Beep statement plays the system's default beep sound.
' It is a simple, non-visual way to provide audible feedback to users.
'
' Syntax:
'   Beep
'
' Behavior:
' - Plays the system-defined default beep sound.
' - Does not pause code execution (non-blocking).
' - Sound depends on user's system sound settings.
'
' Usage scenarios:
' - Notify user that a task has completed (especially long-running macros).
' - Provide subtle alert before/alongside a MsgBox.
' - Minimalist feedback without showing dialogs.
'
' Example 1 (basic use):
'     Beep
'
' Example 2 (beep + message):
'     Beep
'     MsgBox "Export complete!"
'
' Example 3 (invalid input warning):
'     If Not IsNumeric(userInput) Then
'         Beep
'         MsgBox "Please enter a valid number."
'     End If
'
' Limitations:
' - No user interaction (audio only).
' - No control over sound customization (frequency, duration, etc.).
' - Dependent on system sound settings (may be muted or inaudible on some systems).
'
' Notes:
' - For customized sounds or tones, use Windows API calls (not native VBA).
'
' -----------------------------------------
