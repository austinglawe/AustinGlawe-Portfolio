' -----------------------------------------
' VBA Arrays_Collections_Dictionaries:
' Collections
' -----------------------------------------
'
' Purpose:
' - Store groups of items with optional keys.
'
' Create Collection:
'   Dim coll As Collection
'   Set coll = New Collection
'
' Add items:
'   coll.Add "Apple"
'   coll.Add "Banana", "B"
'
' Retrieve items:
'   By index: coll(1)
'   By key: coll("B")
'
' Remove items:
'   coll.Remove 1
'   coll.Remove "B"
'
' Loop through collection:
'   Dim item As Variant
'   For Each item In coll
'       Debug.Print item
'   Next item
'
' Notes:
' - Collection indices start at 1.
' - Supports any data type including objects.
' - No built-in sort method.
'
' Best practices:
' - Use keys for clearer access.
' - Use For Each for iteration.
'
' -----------------------------------------
