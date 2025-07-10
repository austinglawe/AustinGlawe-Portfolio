' -----------------------------------------
' UserForm overview and basics
' -----------------------------------------
'
' What is a UserForm?
' - A UserForm is a custom dialog window you can design in VBA.
' - It allows complex, multi-field input interfaces, supporting a variety of controls.
'
' How to insert:
' - Open the VBA editor.
' - In Project Explorer, right-click on your project.
' - Choose Insert > UserForm.
' - The form will appear as "UserForm1" by default.
'
' Best practice:
' - Rename the UserForm immediately for clarity (for example: CustomerInputForm).
'
' How to show:
'   UserForm1.Show
'   ' Or if renamed:
'   CustomerInputForm.Show
'
' Basic behavior:
' - By default, UserForms show modally (blocking Excel interaction until closed).
' - The form itself acts as a container for controls (Labels, TextBoxes, Buttons, etc.).
'
' Usage scenarios:
' - When multiple inputs need to be collected at once.
' - When custom layout, validation, or appearance is required.
'
' Notes:
' - UserForms have full event handling support (Initialize, Activate, QueryClose, etc.).
' - They can be shown modal or modeless (covered later).
'
' -----------------------------------------
