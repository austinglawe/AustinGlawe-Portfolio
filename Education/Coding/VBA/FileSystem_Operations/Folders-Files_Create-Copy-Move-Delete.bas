' -----------------------------------------
' VBA FileSystem_Operations:
' Creating, copying, moving, deleting files and folders using FSO
' -----------------------------------------
'
' Creating a file:
'   Dim fso As Object
'   Set fso = CreateObject("Scripting.FileSystemObject")
'   Dim file As Object
'   Set file = fso.CreateTextFile("C:\Test\NewFile.txt", True)
'   file.WriteLine "Hello, World!"
'   file.Close
'
' Copying a file:
'   If fso.FileExists("C:\Test\NewFile.txt") Then
'       fso.CopyFile "C:\Test\NewFile.txt", "C:\Test\CopyOfNewFile.txt", True
'   End If
'
' Moving (renaming) a file:
'   If fso.FileExists("C:\Test\CopyOfNewFile.txt") Then
'       fso.MoveFile "C:\Test\CopyOfNewFile.txt", "C:\Test\MovedFile.txt"
'   End If
'
' Deleting a file:
'   If fso.FileExists("C:\Test\MovedFile.txt") Then
'       fso.DeleteFile "C:\Test\MovedFile.txt"
'   End If
'
' Creating a folder:
'   If Not fso.FolderExists("C:\Test\NewFolder") Then
'       fso.CreateFolder "C:\Test\NewFolder"
'   End If
'
' Moving (renaming) a folder:
'   If fso.FolderExists("C:\Test\NewFolder") Then
'       fso.MoveFolder "C:\Test\NewFolder", "C:\Test\MovedFolder"
'   End If
'
' Deleting a folder:
'   If fso.FolderExists("C:\Test\MovedFolder") Then
'       fso.DeleteFolder "C:\Test\MovedFolder"
'   End If
'
' Notes:
' - Always check existence before operations.
' - FSO does not support direct folder copying.
' - Use error handling for permission and locking issues.
'
' -----------------------------------------
