' -----------------------------------------
' UserForm data handling overview
' -----------------------------------------
'
' Reading data from controls:
' - TextBox: TextBox1.Value or TextBox1.Text
' - CheckBox: CheckBox1.Value (True/False)
' - ComboBox: ComboBox1.Value (selected value)
' - ListBox: ListBox1.Value (selected item, single-select)
'
' Writing data to controls:
' - Populate controls before showing form:
'     TextBox1.Value = "John Doe"
'     CheckBox1.Value = True
'     ComboBox1.Value = "Option 1"
'
' Handling multi-select ListBox:
' - Iterate through .Selected and .List properties:
'     Dim i As Long
'     For i = 0 To ListBox1.ListCount - 1
'         If ListBox1.Selected(i) Then
'             Debug.Print "Selected: " & ListBox1.List(i)
'         End If
'     Next i
'
' Writing form data to worksheet:
' - Example:
'     Private Sub SubmitButton_Click()
'         Sheets("Sheet1").Range("A1").Value = TextBox1.Value
'         Sheets("Sheet1").Range("A2").Value = CheckBox1.Value
'         Sheets("Sheet1").Range("A3").Value = ComboBox1.Value
'         Unload Me
'     End Sub
'
' Handling Cancel scenarios:
' - Respect user cancel action:
'     Private cancelled As Boolean
'
'     Private Sub CancelButton_Click()
'         cancelled = True
'         Unload Me
'     End Sub
'
'     ' In calling code:
'     MyForm.Show
'     If Not cancelled Then
'         ' Save form data
'     End If
'
' Best practices:
' - Validate inputs before saving (e.g., required fields, numeric checks).
' - Initialize fields in UserForm_Initialize.
' - Respect Cancel / close without saving.
'
' -----------------------------------------

