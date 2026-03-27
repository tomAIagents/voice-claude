' ── Экземпляр 3 ─────────────────────────────────────────────────────────
Const INSTANCE_NUM  = 3
Const CONFIG_FILE   = "config_3.json"
' ─────────────────────────────────────────────────────────────────────────────

Dim fso, dir, pythonw, shell
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

dir = fso.GetParentFolderName(WScript.ScriptFullName)

Dim candidates(4)
candidates(0) = shell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Programs\Python\Python311\pythonw.exe"
candidates(1) = shell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Programs\Python\Python312\pythonw.exe"
candidates(2) = shell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Programs\Python\Python310\pythonw.exe"
candidates(3) = "C:\Python311\pythonw.exe"
candidates(4) = "C:\Python312\pythonw.exe"

pythonw = ""
Dim i
For i = 0 To 4
    If fso.FileExists(candidates(i)) Then
        pythonw = candidates(i)
        Exit For
    End If
Next

If pythonw = "" Then
    MsgBox "Python не найден. Запустите install.bat.", 16, "VoiceClaude #3"
    WScript.Quit
End If

Dim psCmd
psCmd = "Start-Process '" & pythonw & "' -ArgumentList '" & dir & "\main.py --instance 3 --config config_3.json' -Verb RunAs"
shell.Run "powershell -WindowStyle Hidden -Command " & Chr(34) & psCmd & Chr(34), 0, False
