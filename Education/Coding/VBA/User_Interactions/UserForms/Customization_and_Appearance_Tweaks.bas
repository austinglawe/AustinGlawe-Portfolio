' -----------------------------------------
' UserForm customization and appearance tips
' -----------------------------------------
'
' Background and foreground colors:
' - BackColor: sets background color of form or control.
'     Example:
'         UserForm1.BackColor = RGB(240, 240, 240)
'
' - ForeColor: sets text color for supported controls.
'
' Font customization:
' - UserForm or individual controls can have customized fonts:
'     UserForm1.Font.Name = "Arial"
'     UserForm1.Font.Size = 10
'     Label1.Font.Bold = True
'
' Borders and special effects:
' - BorderStyle (for controls like TextBox, Frame):
'     TextBox1.BorderStyle = fmBorderStyleSingle
'
' - SpecialEffect:
'     Frame1.SpecialEffect = fmSpecialEffectSunken
'     Options include Flat, Sunken, Raised.
'
' Default and cancel buttons:
' - Default = True allows Enter key to trigger a button.
' - Cancel = True allows Esc key to trigger a button.
'
' Preventing resizing or hiding title bar:
' - No native VBA property to hide title bar or close button.
' - Requires advanced Windows API calls (not covered here).
'
' Appearance best practices:
' - Maintain consistent fonts and colors for a clean look.
' - Group related controls visually (Frames recommended).
' - Use alignment tools for orderly layout.
' - Avoid clutter: show only essential controls.
'
' Notes:
' - UserForms do not support automatic resizing: manual layout adjustment required for responsiveness.
'
' -----------------------------------------

