' -----------------------------------------
' VBA UserForms_Basics:
' Common UserForm controls and handling
' -----------------------------------------
'
' Common controls:
' - TextBox (txt): text input
' - Label (lbl): static text
' - CommandButton (cmd): triggers actions
' - ComboBox (cbo): dropdown list
' - ListBox (lst): multi-selection list
' - CheckBox (chk): true/false toggle
' - OptionButton (opt): radio button group
'
' Handling events:
'   Private Sub cmdClose_Click()
'       Unload Me
'   End Sub
'
' Populate ComboBox/ListBox:
'   Private Sub UserForm_Initialize()
'       Me.cboOptions.AddItem "Option 1"
'       Me.cboOptions.AddItem "Option 2"
'       Me.lstChoices.AddItem "Choice A"
'       Me.lstChoices.AddItem "Choice B"
'   End Sub
'
' Reading values:
'   Dim userName As String
'   userName = Me.txtName.Value
'
' Validating input:
'   Private Sub cmdSubmit_Click()
'       If Me.txtName.Value = "" Then
'           MsgBox "Please enter a name."
'           Me.txtName.SetFocus
'           Exit Sub
'       End If
'       Unload Me
'   End Sub
'
' Best practices:
' - Use clear naming prefixes.
' - Populate controls in Initialize.
' - Validate inputs before processing.
' - Provide feedback to users.
'
' -----------------------------------------
