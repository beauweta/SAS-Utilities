'https://stackoverflow.com/questions/4200028/vbscript-list-all-pdf-files-in-folder-and-subfolders

Wscript.Echo "begin."
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objSuperFolder = objFSO.GetFolder(WScript.Arguments(0))
Call ShowSubfolders (objSuperFolder)

Wscript.Echo "end."

WScript.Quit 0

Sub ShowSubFolders(fFolder)
    Set objFolder = objFSO.GetFolder(fFolder.Path)
    Set colFiles = objFolder.Files
    For Each objFile in colFiles
        If UCase(objFSO.GetExtensionName(objFile.name)) = "PDF" Then
            Wscript.Echo objFile.Name
        End If
    Next

    For Each Subfolder in fFolder.SubFolders
        ShowSubFolders(Subfolder)
    Next
End Sub
