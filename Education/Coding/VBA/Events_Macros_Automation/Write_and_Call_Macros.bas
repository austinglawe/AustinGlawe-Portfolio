' -----------------------------------------
' VBA Events_Macros_Automation:
' Writing and calling macros
' -----------------------------------------
'
' Purpose:
' - Automate tasks via Sub procedures (macros).
'
' Writing a simple macro:
'   Sub HelloWorld()
'       MsgBox "Hello, world!"
'   End Sub
'
' Calling macros from other procedures:
'   Sub CallHello()
'       Call HelloWorld
'       ' Or simply:
'       HelloWorld
'   End Sub
'
' Running macros manually:
' - Developer tab > Macros > Select and Run
' - Assign to buttons, shapes, or keyboard shortcuts
'
' Notes on parameters:
' - Macros run from dialog must have no parameters.
' - Use parameterized Subs only when called from code.
'
' Best practices:
' - Use clear descriptive names.
' - Comment code.
' - Include error handling for robustness.
'
' -----------------------------------------
