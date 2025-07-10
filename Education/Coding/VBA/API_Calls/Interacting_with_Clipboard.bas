' -----------------------------------------
' VBA API_Calls:
' Interacting with the system clipboard
' -----------------------------------------
'
' Common API functions:
'   OpenClipboard, CloseClipboard
'   EmptyClipboard
'   GetClipboardData, SetClipboardData
'
' Declarations:
' #If VBA7 Then
'   Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
'   Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
'   Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
'   Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As LongPtr
'   Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
' #Else
'   Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
'   Declare Function CloseClipboard Lib "user32" () As Long
'   Declare Function EmptyClipboard Lib "user32" () As Long
'   Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
'   Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
' #End If
'
' Clipboard format constants:
'   Const CF_TEXT = 1
'   Const CF_BITMAP = 2
'   Const CF_UNICODETEXT = 13
'
' Best practices:
' - Always open and close clipboard properly.
' - Use correct data formats.
' - Handle errors gracefully.
' - Be cautious to avoid interfering with user clipboard.
'
' -----------------------------------------
