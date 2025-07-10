' -----------------------------------------
' Application.InputBox overview and guidance
' -----------------------------------------
'
' Application.InputBox is an enhanced version of InputBox.
' It adds the ability to specify the expected data type and enables range selection directly from the worksheet UI.
'
' Syntax:
'   Application.InputBox(prompt, [title], [default], [left], [top], [helpfile], [context], [type])
'
' Parameters:
' - prompt (required): Message to display.
' - title (optional): Window title text (defaults to "Microsoft Excel").
' - default (optional): Pre-filled value.
' - left, top (optional): Screen position (points, not twips).
' - helpfile, context (optional): Legacy help system support (rarely used).
' - type (optional): Expected input type (enforces basic validation).
'
' Type values:
'   0 – Formula (default)
'   1 – Number (numeric input)
'   2 – String (text input)
'   4 – Boolean (True/False input)
'   8 – Range (enables user to click/select worksheet cells)
'   16 – Error value
'   64 – Array
'
' Return value:
' - If user clicks OK: returns input as native VBA type (number, string, range, etc.).
' - If user clicks Cancel: returns False (not empty string!).
'
' Usage scenarios:
' - When you want to ensure the correct input type (e.g., number).
' - When you want the user to visually select a range in the worksheet.
'
' Example 1 (number input):
'     Dim userNumber As Variant
'     userNumber = Application.InputBox("Enter a number:", "Number Prompt", Type:=1)
'     If userNumber = False Then
'         MsgBox "User clicked Cancel."
'     Else
'         MsgBox "You entered: " & userNumber
'     End If
'
' Example 2 (text input):
'     Dim userText As Variant
'     userText = Application.InputBox("Enter some text:", "Text Prompt", Type:=2)
'
' Example 3 (range selection):
'     Dim userRange As Range
'     Dim result As Variant
'     result = Application.InputBox("Select a range:", "Range Prompt", Type:=8)
'     If result Is Nothing Or result = False Then
'         MsgBox "No range selected."
'     Else
'         Set userRange = result
'         MsgBox "You selected: " & userRange.Address
'     End If
'
' Notes:
' - Unlike InputBox, Application.InputBox returns native types.
' - Must check explicitly for Cancel (returns False, not empty string).
' - Great for typed input validation and range picking.
'
' -----------------------------------------
