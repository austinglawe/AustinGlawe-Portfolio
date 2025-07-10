' -----------------------------------------
' VBA Arrays_Collections_Dictionaries:
' Dictionaries (Scripting.Dictionary)
' -----------------------------------------
'
' Purpose:
' - Store key-value pairs with fast lookup.
'
' Create Dictionary (late binding):
'   Dim dict As Object
'   Set dict = CreateObject("Scripting.Dictionary")
'
' Add items:
'   dict.Add "A", "Apple"
'   dict.Add "B", "Banana"
'
' Access items:
'   Debug.Print dict("A")
'
' Check if key exists:
'   If dict.Exists("B") Then MsgBox "Key B is present."
'
' Remove items:
'   dict.Remove "A"
'
' Loop through keys and items:
'   Dim key As Variant
'   For Each key In dict.Keys
'       Debug.Print key & ": " & dict(key)
'   Next key
'
' Useful properties/methods:
' - .Count
' - .Exists(key)
' - .Keys
' - .Items
' - .Remove(key)
' - .RemoveAll
'
' Best practices:
' - Use for fast key-based lookup.
' - Use late binding for portability (no reference needed).
' - Add reference for early binding (Microsoft Scripting Runtime).
'
' -----------------------------------------
