' -----------------------------------------
' VBA File_Handling:
' FileSystemObject file manipulation
' -----------------------------------------
'
' Purpose:
' - Read, write, delete files using FileSystemObject.
'
' Initialization:
'   Dim fso As Object
'   Set fso = CreateObject("Scripting.FileSystemObject")
'
' Reading text file:
'   Dim ts As Object
'   Set ts = fso.OpenTextFile("C:\Test\sample.txt", 1)  ' 1 = ForReading
'
'   Do While Not ts.AtEndOfStream
'       Debug.Print ts.ReadLine
'   Loop
'
'   ts.Close
'
' Writing new file (overwrite if exists):
'   Set ts = fso.CreateTextFile("C:\Test\output.txt", True)
'   ts.WriteLine "Line 1"
'   ts.WriteLine "Line 2"
'   ts.Close
'
' Appending to file:
'   Set ts = fso.OpenTextFile("C:\Test\output.txt", 8)  ' 8 = ForAppending
'   ts.WriteLine "Appended line"
'   ts.Close
'
' Deleting file:
'   If fso.FileExists("C:\Test\output.txt") Then
'       fso.DeleteFile "C:\Test\output.txt"
'   End If
'
' Constants for IOMode:
' - ForReading = 1
' - ForWriting = 2
' - ForAppending = 8
'
' Best practices:
' - Always close TextStream objects.
' - Use FileExists before deleting.
' - CreateTextFile for clean overwrite or new file.
' - OpenTextFile for appending.
'
' -----------------------------------------
