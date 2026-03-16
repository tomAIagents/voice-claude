Dim dir
dir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "powershell -WindowStyle Hidden -Command ""Start-Process pythonw -ArgumentList '" & dir & "\main.py' -Verb RunAs""", 0, False
