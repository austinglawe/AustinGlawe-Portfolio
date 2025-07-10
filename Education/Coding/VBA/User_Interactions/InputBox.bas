' -----------------------------------------
' InputBox overview and guidance
' -----------------------------------------
'
' The InputBox function prompts the user for a simple input string.
' It is modal (halts execution until user responds) and intended for
' lightweight single-value prompts.
'
' Syntax:
'   InputBox(prompt, [title], [default], [xpos], [ypos], [helpfile], [context])
'
' Parameters:
' - prompt (required): Message displayed to the user.
' - title (optional): Window title text (defaults to application name).
' - default (optional): Pre-filled value in text box.
' - xpos, ypos (optional): Position of window on screen (twips).
' - helpfile, context (optional): Legacy help system support (rarely used).
'
' Behavior:
' - Always returns a String.
' - If user clicks OK: returns what they typed.
' - If user clicks Cancel: returns empty string "".
'
' Usage scenarios:
' - Quick prompts for single text or numeric entry.
' - Lightweight interaction without requiring a UserForm.
'
' Example 1 (basic):
'     Dim userName As String
'     userName = InputBox("Enter your name:", "Name Prompt")
'
' Example 2 (with default):
'     Dim userName As String
'     userName = InputBox("Enter your name:", "Name Prompt", "John Doe")
'
' Example 3 (handling Cancel):
'     Dim userName As String
'     userName = InputBox("Enter your name:", "Name Prompt", "John Doe")
'     If userName = "" Then
'         MsgBox "No name entered (or user clicked Cancel)."
'     Else
'         MsgBox "Hello, " & userName
'     End If
'
' Limitations:
' - No input validation.
' - Only one value can be requested.
' - Always returns a String; convert manually if numeric input is required.
'
' Alternatives:
' - For input validation or typed input, see Application.InputBox.
'
' -----------------------------------------
