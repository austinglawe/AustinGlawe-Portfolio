' -----------------------------------------
' VBA FileSystem_Operations:
' Reading and writing file contents
' -----------------------------------------
'
' Reading a text file:
'   Dim fso As Object, txtStream As Object
'   Set fso = CreateObject("Scripting.FileSystemObject")
'   Set txtStream = fso.OpenTextFile("C:\Test\Sample.txt", 1)  ' 1=ForReading
'   Dim content As String
'   content = txtStream.ReadAll
'   txtStream.Close
'
' Writing to a text file:
'   Set txtStream = fso.CreateTextFile("C:\Test\Output.txt", True)  ' True=overwrite
'   txtStream.WriteLine "Hello, world!"
'   txtStream.Close
'
' Appending to a text file:
'   Set txtStream = fso.OpenTextFile("C:\Test\Output.txt", 8, True)  ' 8=ForAppending
'   txtStream.WriteLine "Additional line."
'   txtStream.Close
'
' Reading binary files with ADODB.Stream:
'   Dim stream As Object
'   Set stream = CreateObject("ADODB.Stream")
'   stream.Type = 1  ' Binary
'   stream.Open
'   stream.LoadFromFile "C:\Test\image.jpg"
'   Dim byteData() As Byte
'   byteData = stream.Read
'   stream.Close
'
' Best practices:
' - Use text streams for text files.
' - Use ADODB.Stream for binary files.
' - Always close streams after use.
' - Check file existence before reading.
'
' -----------------------------------------
