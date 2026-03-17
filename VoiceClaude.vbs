Dim fso, dir, pythonw, shell
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

dir = fso.GetParentFolderName(WScript.ScriptFullName)

pythonw = ""

' 1. Ищем pythonw через where, пропуская WindowsApps (заглушка Microsoft Store)
On Error Resume Next
Dim exec
Set exec = shell.Exec("cmd /c where pythonw 2>nul")
If Err.Number = 0 Then
    Do While Not exec.StdOut.AtEndOfStream
        Dim line
        line = exec.StdOut.ReadLine()
        If InStr(LCase(line), "windowsapps") = 0 Then
            If fso.FileExists(line) Then
                pythonw = line
                Exit Do
            End If
        End If
    Loop
End If
On Error GoTo 0

' 2. Если не нашли через where — проверяем стандартные пути
If pythonw = "" Then
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

    Dim i
    For i = 0 To 9
        If fso.FileExists(candidates(i)) Then
            pythonw = candidates(i)
            Exit For
        End If
    Next
End If

If pythonw = "" Then
    MsgBox "Python not found." & vbCrLf & vbCrLf & "Please run install.bat as Administrator.", 16, "VoiceClaude"
    WScript.Quit
End If

Dim mainScript
mainScript = dir & "\main.py"

shell.Run "powershell -WindowStyle Hidden -Command ""Start-Process '"& pythonw &"' -ArgumentList '""""" & mainScript & """""' -Verb RunAs""", 0, False
