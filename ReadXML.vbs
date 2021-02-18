'https://stackoverflow.com/questions/30224180/how-to-traverse-subfolders-in-a-zip-file-and-unzip-files-with-specific-extension

Set fso = CreateObject("Scripting.FileSystemObject")
Set app = CreateObject("Shell.Application")

Sub ExtractByExtension(fldr, ext, dst)
  For Each f In fldr.Items
    If f.Type = "File folder" Then
      ExtractByExtension f.GetFolder, ext, dst
    ElseIf LCase(fso.GetExtensionName(f.Name)) = LCase(ext) Then
      app.NameSpace(dst).CopyHere f.Path
    End If
  Next
End Sub

ExtractByExtension app.NameSpace("C:\path\to\your.zip"), "txt", "C:\output"
