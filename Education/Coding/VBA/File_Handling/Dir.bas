' -----------------------------------------
' VBA File_Handling: Dir function overview
' -----------------------------------------
'
' Purpose:
' - Lightweight file/folder existence checks.
' - Iterate through files/folders in a directory.
'
' Syntax:
'   result = Dir([path], [attributes])
'
' Attributes:
' - vbNormal: Normal files (default)
' - vbDirectory: Include directories
' - vbHidden: Include hidden files
' - vbReadOnly: Include read-only files
' - vbSystem: Include system files
' - Combine attributes with Or.
'
' Common usage:
'
' 1. Check if file exists:
'   If Dir("C:\Test\MyFile.txt") <> "" Then
'       MsgBox "File exists!"
'   End If
'
' 2. Check if folder exists:
'   If Dir("C:\Test", vbDirectory) <> "" Then
'       MsgBox "Folder exists!"
'   End If
'
' 3. Iterate all *.txt files:
'   Dim fileName As String
'   fileName = Dir("C:\Test\*.txt")
'   Do While fileName <> ""
'       Debug.Print fileName
'       fileName = Dir
'   Loop
'
' 4. Iterate folders:
'   Dim folderName As String
'   folderName = Dir("C:\Test\*", vbDirectory)
'   Do While folderName <> ""
'       If folderName <> "." And folderName <> ".." Then
'           If (GetAttr("C:\Test\" & folderName) And vbDirectory) = vbDirectory Then
'               Debug.Print "Folder: " & folderName
'           End If
'       End If
'       folderName = Dir
'   Loop
'
' Notes:
' - First call initializes sequence, subsequent Dir calls continue iteration.
' - Returns "" when done.
' - Resets if a new path is passed mid-loop.
'
' Best practices:
' - Quick existence checks: use Dir.
' - Complex scenarios: prefer FileSystemObject.
'
' -----------------------------------------
