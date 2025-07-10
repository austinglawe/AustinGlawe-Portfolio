' -----------------------------------------
' VBA Loops_Conditionals:
' Select Case structure
' -----------------------------------------
'
' Purpose:
' - Clean alternative to multiple If/ElseIf checks when evaluating one expression.
'
' Syntax:
'   Select Case expression
'       Case value1
'           ' Code for value1
'       Case value2
'           ' Code for value2
'       Case Else
'           ' Code if no match
'   End Select
'
' Example:
'   Dim score As Integer
'   score = 85
'
'   Select Case score
'       Case Is >= 90
'           MsgBox "A"
'       Case Is >= 80
'           MsgBox "B"
'       Case Is >= 70
'           MsgBox "C"
'       Case Else
'           MsgBox "F"
'   End Select
'
' Multiple values in a case:
'   Select Case fruit
'       Case "Apple", "Orange", "Banana"
'           MsgBox "Known fruit"
'       Case Else
'           MsgBox "Unknown fruit"
'   End Select
'
' Range of values:
'   Select Case score
'       Case 0 To 59
'           MsgBox "Fail"
'       Case 60 To 79
'           MsgBox "Pass"
'       Case 80 To 100
'           MsgBox "Excellent"
'       Case Else
'           MsgBox "Invalid"
'   End Select
'
' Best practices:
' - Use when testing a single expression against many values.
' - Include Case Else as a safety net.
' - Improves readability vs. If/ElseIf chains.
'
' -----------------------------------------
