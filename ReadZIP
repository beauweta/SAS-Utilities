'https://stackoverflow.com/questions/30224180/how-to-traverse-subfolders-in-a-zip-file-and-unzip-files-with-specific-extension

' *** INIT ***
Set objApp = CreateObject("Shell.Application")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' *** INPUT PARAMETER: FOLDER TO READ ***
'Set objTopFolder = objFSO.GetFolder(WScript.Arguments(0))
Set objTopFolder = objFSO.GetFolder("\\retailproducts\RetailProducts\Retail Products\Portfolio Management\CCCFA Fee Review (legally privileged)")

' *** OUTPUT PARAMETER: FILE TO WRITE ***
Set objLog = objFSO.CreateTextFile("out9.log", true)

' *** RUN SCRIPT ***
Wscript.Echo "begin."
Call ShowSubfolders (objTopFolder)
objLog.close
Wscript.Echo "end."
WScript.Quit 0

' *** FIND ALL THE MS OFFICE XML-FORMAT FILES 
' *** WHEN FOUND: ADD .ZIP EXTENSION
' ***             CALL ROUTINE TO EXTACT THE AUTHORS
' ***             REMOVE .ZIP EXTENSION
Sub ShowSubFolders(fFolder)
    Set objFolder = objFSO.GetFolder(fFolder.Path)
    Set colFiles = objFolder.Files
	On Error Resume Next

    For Each objFile in colFiles
	   Do
       If Left(objFSO.GetFileName(objFile.name),1) = "~" Then Exit Do

       If UCase(objFSO.GetExtensionName(objFile.name)) = "XLSX" _ 
       or UCase(objFSO.GetExtensionName(objFile.name)) = "XLSB" _ 
       or UCase(objFSO.GetExtensionName(objFile.name)) = "XLSM" _ 
       or UCase(objFSO.GetExtensionName(objFile.name)) = "DOCX" _ 
       or UCase(objFSO.GetExtensionName(objFile.name)) = "PPTX" Then
			objLog.WriteLine (objFile.path + vbCrLf  )
			FilePath  =objFile.Path
			FilePath2 =objFile.Path+".ZIP"
			objFSO.MoveFile FilePath, FilePath2
			ExtractByExtension objApp.NameSpace(FilePath2)
			objFSO.MoveFile FilePath2,FilePath  
		End if
        If UCase(objFSO.GetExtensionName(objFile.name)) = "XLS" _ 
        or UCase(objFSO.GetExtensionName(objFile.name)) = "DOC" _ 
        or UCase(objFSO.GetExtensionName(objFile.name)) = "PPT" Then 
			objLog.WriteLine (objFile.path + " / NODATA " + vbCrLf  )
		End if
		Loop While False
    Next

	For Each Subfolder in fFolder.SubFolders
		ShowSubFolders(Subfolder)
	Next
End Sub

' *** SCAN THE ZIP FILE, FIND THE core.xml FILE
' *** COPY core.xml  OUT AND SAVE ITS CONTENTS
Sub ExtractByExtension(fldr)
    tmpFolder="c:\temp\a\"
	On Error Resume Next
	For Each f In fldr.Items
	If f.Type = "File folder" Then
		ExtractByExtension f.GetFolder
	ElseIf f.name="core.xml" Then 
		If objFSO.FileExists (tmpFolder + "core.xml") Then
			objFSO.DeleteFile(tmpFolder + "core.xml")
		End if
		objApp.NameSpace(tmpFolder).CopyHere f.Path
		set objFileToRead = objFSO.OpenTextFile(tmpFolder + "core.xml",1)
		strFileText = objFileToRead.ReadAll()
		objFileToRead.Close
		objLog.WriteLine ( " => " + strFileText + vbCrLf )
		strFileText=" "
		If objFSO.FileExists (tmpFolder + "core.xml") Then
			objFSO.DeleteFile(tmpFolder + "core.xml")
		End If
	End If
  Next
End Sub

