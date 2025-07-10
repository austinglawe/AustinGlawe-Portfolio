' -----------------------------------------
' VBA FileSystem_Operations:
' Accessing file and folder metadata
' -----------------------------------------
'
' File properties:
' - .Name, .Path, .Size (bytes)
' - .DateCreated, .DateLastModified, .DateLastAccessed
' - .Attributes (bitmask)
'
' Example file metadata:
'   Set file = fso.GetFile("C:\Test\Example.txt")
'   Debug.Print file.Name, file.Size, file.DateCreated
'
' Folder properties:
' - .Name, .Path
' - .DateCreated, .DateLastModified
' - .Attributes
'
' Example folder metadata:
'   Set folder = fso.GetFolder("C:\Test")
'   Debug.Print folder.Name, folder.DateCreated
'
' Attributes bitmask examples:
' - 1 = Read-only
' - 2 = Hidden
' - 4 = System
' - 16 = Directory (folder)
' - 32 = Archive
' - Check attributes with bitwise AND:
'     If (file.Attributes And 2) <> 0 Then Debug.Print "Hidden"
'
' Best practices:
' - Use metadata for filtering and sorting.
' - Calculate folder size by summing files recursively.
' - Use bitwise operations for attributes.
'
' -----------------------------------------
