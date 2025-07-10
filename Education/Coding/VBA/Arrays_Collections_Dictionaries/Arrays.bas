' -----------------------------------------
' VBA Arrays_Collections_Dictionaries:
' Arrays
' -----------------------------------------
'
' Purpose:
' - Store multiple values of the same type indexed by number.
'
' Static arrays (fixed size):
'   Dim arr(1 To 5) As Integer
'
' Dynamic arrays (resizable):
'   Dim arr() As Integer
'   ReDim arr(1 To 10)
'   ReDim Preserve arr(1 To 15)  ' Preserve existing values
'
' Multidimensional arrays:
'   Dim arr(1 To 3, 1 To 2) As String
'   arr(1,1) = "Hello"
'   arr(3,2) = "World"
'
' Access bounds:
'   LBound(arr), UBound(arr)
'
' Example:
'   Dim arr() As String
'   ReDim arr(1 To 3)
'   arr(1) = "Apple"
'   arr(2) = "Banana"
'   arr(3) = "Cherry"
'
' Best practices:
' - Use LBound and UBound for loops.
' - Use dynamic arrays for flexibility.
' - Use ReDim Preserve cautiously (performance).
'
' -----------------------------------------
