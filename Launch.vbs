Dim objShell, objFSO, strDir
Set objShell = CreateObject("WScript.Shell")
Set objFSO   = CreateObject("Scripting.FileSystemObject")
strDir = objFSO.GetParentFolderName(WScript.ScriptFullName)
objShell.Run "powershell -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & strDir & "\Windrose-Server-Manager.ps1""", 0, False
