' -----------------------------------------
' VBA UserForms_Basics:
' Passing data between UserForms and worksheets
' -----------------------------------------
'
' Writing data to worksheet:
'   Private Sub cmdSubmit_Click()
'       Worksheets("Sheet1").Range("A1").Value = Me.txtName.Value
'       Unload Me
'   End Sub
'
' Reading data from worksheet:
'   Private Sub UserForm_Initialize()
'       Me.txtName.Value = Worksheets("Sheet1").Range("A1").Value
'   End Sub
'
' Writing multiple values:
'   For i = 1 To 5
'       Worksheets("Sheet1").Cells(i, 1).Value = Me.Controls("txtValue" & i).Value
'   Next i
'
' Reading multiple values:
'   For i = 1 To 5
'       Me.Controls("txtValue" & i).Value = Worksheets("Sheet1").Cells(i, 1).Value
'   Next i
'
' Best practices:
' - Use consistent control naming.
' - Validate data before writing.
' - Use Initialize to load data, buttons to save.
' - Handle empty/invalid cells gracefully.
'
' -----------------------------------------
