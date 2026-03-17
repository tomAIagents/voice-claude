Dim fso, dir, shell
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

dir = fso.GetParentFolderName(WScript.ScriptFullName)

Dim setupScript
setupScript = dir & "\setup.ps1"

shell.Run "powershell -ExecutionPolicy Bypass -File """ & setupScript & """", 1, False
