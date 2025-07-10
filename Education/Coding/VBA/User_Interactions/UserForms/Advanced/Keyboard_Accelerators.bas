' -----------------------------------------
' Advanced UserForm customization:
' Keyboard accelerators (Alt + key shortcuts)
' -----------------------------------------
'
' Purpose:
' - Improve keyboard accessibility by allowing Alt + key shortcuts for controls.
'
' How to define:
' - Insert "&" before the letter in the Caption property you want as the accelerator.
'
' Example:
'   CommandButton1.Caption = "&Submit"
'   ' Displays as "Submit" with "S" underlined.
'   ' User can press Alt + S to activate the button.
'
' Supported controls:
' - CommandButton
' - Label
' - CheckBox
' - OptionButton
' - Frame
'
' Notes:
' - If multiple controls share the same accelerator key, pressing Alt + key cycles between them.
' - TabIndex determines focus order for cycling.
' - To display a literal "&", use "&&" in Caption (e.g., "Save && Close").
'
' Best practices:
' - Define accelerators for all primary actions (Submit, Cancel, etc.).
' - Choose mnemonic letters that are easy to remember and avoid conflicts.
' - Ensure logical TabIndex order for usability.
'
' -----------------------------------------
