' -----------------------------------------
' Worksheet-based user prompting overview
' -----------------------------------------
'
' This is not a VBA dialog or function, but a design technique:
' guiding users directly in the worksheet using cell formatting,
' selection, prompts, and comments.
'
' Common techniques:
'
' 1. Selecting or activating a cell or range:
'     Range("A1").Select
'
' 2. Highlighting cells to indicate where user should edit:
'     With Range("A1:A5")
'         .Interior.Color = RGB(255, 255, 200) ' light yellow background
'         .Font.Bold = True
'     End With
'
' 3. Displaying text prompts inside cells:
'     Range("A1").Value = "Please enter your name below:"
'
' 4. Using comments to add guidance:
'     Range("B2").AddComment "Enter your date of birth here."
'
' 5. Combining with StatusBar and Application.InputBox:
'     Range("A1:A10").Select
'     Application.StatusBar = "Please review and edit cells A1:A10"
'
' Usage scenarios:
' - When you want the user to work naturally in the sheet itself.
' - When designing worksheets that behave like templates or forms.
' - When intrusive popups are unnecessary or undesirable.
'
' Best practices:
' - Combine with visual cues (colors, borders).
' - Protect areas of the sheet that should not be edited.
' - Use Data Validation where appropriate for restricting input.
'
' Limitations:
' - No built-in enforcement: user can ignore your visual cues.
' - Less control compared to modal dialogs.
'
' -----------------------------------------
