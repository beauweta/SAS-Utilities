
JobName = "Job1002"    ' No spaces in job name
Status  = "Starting"   ' Starting or Ending

Dim objShell
Set objShell = Wscript.CreateObject("WScript.Shell")
objShell.Run "JournalWrite.vbs " & JobName & " " & Status 
