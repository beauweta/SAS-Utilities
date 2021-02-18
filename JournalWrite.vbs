
Const File       = "C:\temp\file.txt"

Const ForReading = 1
Const ForWriting = 2
Text             = Date() & " " & Time() & " " & WScript.Arguments(0) & " " & WScript.Arguments(1)
Set objFSO       = CreateObject("Scripting.FileSystemObject")
Set objFile      = objFSO.OpenTextFile(File, ForReading)
strContents      = objFile.ReadAll
objFile.Close
Set objFile      = objFSO.OpenTextFile(File, ForWriting)
objFile.WriteLine strContents & Text
objFile.Close
