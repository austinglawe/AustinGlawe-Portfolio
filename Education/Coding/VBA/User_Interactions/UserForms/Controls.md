' -----------------------------------------
' UserForm controls overview
' -----------------------------------------
'
' UserForms act as containers for many types of controls.
' Each control serves a different interaction purpose.
'
' Common controls:
'
' Label
' - Displays static text (instructions, context).
' - Example: "Enter your name:"
'
' TextBox
' - Single-line or multiline text input.
' - Used for accepting free-form text from the user.
'
' CommandButton
' - Clickable button.
' - Triggers actions when clicked (Submit, Cancel, etc.).
'
' CheckBox
' - Toggle control (checked or unchecked).
' - Used for independent True/False choices.
'
' OptionButton
' - Radio button.
' - Allows one selection from a set (group using a Frame).
'
' ComboBox
' - Dropdown list (with optional free text entry).
' - Used for selecting from predefined values or typing new ones.
'
' ListBox
' - Displays a list of items.
' - Allows single or multiple selections.
'
' Frame
' - Groups related controls visually and logically.
' - Especially useful for grouping OptionButtons to ensure correct behavior.
'
' MultiPage
' - Provides tabbed pages within a UserForm.
' - Allows organization of large or complex forms into sections.
'
' Image
' - Displays pictures or logos on the UserForm.
'
' ScrollBar
' - Horizontal or vertical scrollbar control.
' - Adjust numeric values using a slider interface.
'
' SpinButton
' - Up/down arrow control.
' - Used to increment/decrement a numeric value visually.
'
' Best practices:
' - Always label inputs clearly using Labels.
' - Use Frames to organize OptionButtons properly.
' - Use CheckBoxes for independent True/False options.
' - Use ComboBox when you have many options and want to allow typing too.
' - Set Tab Order thoughtfully for good user experience.
'
' Notes:
' - Controls can be added at design time via the Toolbox or at runtime using Controls.Add (covered later).
'
' -----------------------------------------
