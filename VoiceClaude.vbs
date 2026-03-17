Dim fso, dir, pythonw, shell
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

dir = fso.GetParentFolderName(WScript.ScriptFullName)

' Ищем pythonw.exe в стандартных местах
Dim candidates(9)
candidates(0) = shell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Programs\Python\Python313\pythonw.exe"
candidates(1) = shell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Programs\Python\Python312\pythonw.exe"
candidates(2) = shell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Programs\Python\Python311\pythonw.exe"
candidates(3) = shell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Programs\Python\Python310\pythonw.exe"
candidates(4) = "C:\Program Files\Python313\pythonw.exe"
candidates(5) = "C:\Program Files\Python312\pythonw.exe"
candidates(6) = "C:\Program Files\Python311\pythonw.exe"
candidates(7) = "C:\Program Files\Python310\pythonw.exe"
candidates(8) = "C:\Python312\pythonw.exe"
candidates(9) = "C:\Python311\pythonw.exe"

pythonw = ""
Dim i
For i = 0 To 9
    If fso.FileExists(candidates(i)) Then
        pythonw = candidates(i)
        Exit For
    End If
Next

' Если не нашли — ищем в PATH через where
If pythonw = "" Then
    Dim exec, line
    On Error Resume Next
    Set exec = shell.Exec("cmd /c where pythonw 2>nul")
    If Err.Number = 0 Then
        line = exec.StdOut.ReadLine()
        If fso.FileExists(line) Then pythonw = line
    End If
    On Error GoTo 0
End If

If pythonw = "" Then
    MsgBox "Python не найден." & vbCrLf & "Запустите install.bat от имени администратора.", 16, "VoiceClaude"
    WScript.Quit
End If

Dim mainScript
mainScript = dir & "\main.py"

shell.Run "powershell -WindowStyle Hidden -Command ""Start-Process '" & pythonw & "' -ArgumentList '""""" & mainScript & """""' -Verb RunAs""", 0, False
