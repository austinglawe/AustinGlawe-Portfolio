' -----------------------------------------
' Advanced UserForm customization:
' Handling events for dynamically added controls
' -----------------------------------------
'
' Problem:
' - Controls added at runtime using Controls.Add do not automatically support event handlers.
'
' Solution:
' - Use a class module with WithEvents to "wrap" each dynamic control.
'
' Step-by-step:
'
' 1. Create a class module and name it clsButtonHandler.
'
' 2. In clsButtonHandler:
'     Public WithEvents btn As MSForms.CommandButton
'
'     Private Sub btn_Click()
'         MsgBox "Dynamic button clicked! Caption: " & btn.Caption
'     End Sub
'
' 3. In the UserForm code module:
'     Private ButtonHandlers As Collection
'
'     Private Sub UserForm_Initialize()
'         Dim i As Integer
'         Dim ctl As MSForms.CommandButton
'         Dim handler As clsButtonHandler
'
'         Set ButtonHandlers = New Collection
'
'         For i = 1 To 3
'             Set ctl = Me.Controls.Add("Forms.CommandButton.1", "btnDynamic" & i, True)
'             ctl.Caption = "Dynamic " & i
'             ctl.Top = 20 + ((i - 1) * 30)
'             ctl.Left = 20
'             ctl.Width = 100
'
'             Set handler = New clsButtonHandler
'             Set handler.btn = ctl
'
'             ButtonHandlers.Add handler
'         Next i
'     End Sub
'
' Notes:
' - You must retain references to all clsButtonHandler instances (e.g., in a Collection) or they will go out of scope and stop working.
' - WithEvents is only valid inside a class module.
'
' Benefits:
' - Fully event-driven support for runtime-created controls.
'
' -----------------------------------------
