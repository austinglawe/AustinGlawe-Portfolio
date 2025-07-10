' -----------------------------------------
' UserForm showing, hiding, and modal vs modeless overview
' -----------------------------------------
'
' Showing a UserForm:
' - Use .Show method:
'     UserForm1.Show
'
' Modal behavior (default):
' - Blocks interaction with Excel until the form is closed.
' - Macro execution pauses until the form closes.
'
' Modeless behavior:
' - Allows user to continue interacting with Excel while form remains open.
' - Syntax:
'     UserForm1.Show vbModeless
'
' When to use:
' - Modal:
'     Use when immediate user input is required before continuing.
' - Modeless:
'     Use when the form provides tools or utility functions and should coexist with the workbook UI.
'
' Hiding a UserForm:
' - Syntax:
'     UserForm1.Hide
' - Behavior:
'     Hides the form but retains its state in memory.
'     Controls and values remain as they were until unloaded.
'
' Unloading a UserForm:
' - Syntax:
'     Unload UserForm1
' - Behavior:
'     Fully clears the form from memory.
'     User input and control states are reset.
'     Next time shown, Initialize will fire again.
'
' Summary:
' - .Show: Display the form (modal or modeless).
' - .Hide: Temporarily remove from view (state retained).
' - Unload: Fully remove from memory and reset state.
'
' -----------------------------------------

