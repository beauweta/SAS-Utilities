
Const File       = "C:\temp\file.txt"

Const ForReading   = 1
Const ForWriting   = 2
Const ForAppending = 8
Text             = Date() & " " & Time() & " " & WScript.Arguments(0) & " " & WScript.Arguments(1)
Set objFSO       = CreateObject("Scripting.FileSystemObject")
Set objFile      = objFSO.OpenTextFile(File, ForAppending)
objFile.WriteLine Text
objFile.Close
