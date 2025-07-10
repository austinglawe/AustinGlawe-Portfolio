' -----------------------------------------
' VBA Loops_Conditionals:
' If / ElseIf / Else structure
' -----------------------------------------
'
' Purpose:
' - Make decisions in code based on conditions.
'
' Syntax:
' - Basic:
'     If condition Then
'         ' Code
'     End If
'
' - With Else:
'     If condition Then
'         ' Code if True
'     Else
'         ' Code if False
'     End If
'
' - With ElseIf:
'     If condition1 Then
'         ' Code for condition1
'     ElseIf condition2 Then
'         ' Code for condition2
'     Else
'         ' Code if none True
'     End If
'
' Example:
'   Dim score As Integer
'   score = 85
'
'   If score >= 90 Then
'       MsgBox "Grade: A"
'   ElseIf score >= 80 Then
'       MsgBox "Grade: B"
'   ElseIf score >= 70 Then
'       MsgBox "Grade: C"
'   Else
'       MsgBox "Grade: F"
'   End If
'
' Single-line If:
' - If x > 0 Then MsgBox "Positive"
'
' Best practices:
' - Use ElseIf instead of nested Ifs for clarity.
' - Indent properly for readability.
' - Use parentheses for complex conditions.
'
' -----------------------------------------
