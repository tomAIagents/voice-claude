"""
Создаёт следующий по номеру файл запуска VoiceClaude.
Запускать через create_instance.bat — без аргументов.
"""
import os

DIR = os.path.dirname(os.path.abspath(__file__))

VBS_TEMPLATE = r"""' ── Экземпляр {num} ─────────────────────────────────────────────────────────
Const INSTANCE_NUM  = {num}
Const CONFIG_FILE   = "config_{num}.json"
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
    MsgBox "Python не найден. Запустите install.bat.", 16, "VoiceClaude #{num}"
    WScript.Quit
End If

Dim psCmd
psCmd = "Start-Process '" & pythonw & "' -ArgumentList '" & dir & "\main.py --instance {num} --config config_{num}.json' -Verb RunAs"
shell.Run "powershell -WindowStyle Hidden -Command " & Chr(34) & psCmd & Chr(34), 0, False
"""

# Найти следующий свободный номер
num = 1
while os.path.exists(os.path.join(DIR, f"VoiceClaude_{num}.vbs")):
    num += 1

# Записать VBS
vbs_path = os.path.join(DIR, f"VoiceClaude_{num}.vbs")
with open(vbs_path, "w", encoding="utf-8") as f:
    f.write(VBS_TEMPLATE.format(num=num))

print(f"Создан VoiceClaude_{num}.vbs")
