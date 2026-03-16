#Requires AutoHotkey v2.0

; Получаем текст как аргумент командной строки
text := A_Args[1]

if WinExist("ahk_exe Code.exe") {
    WinActivate("ahk_exe Code.exe")
    WinWaitActive("ahk_exe Code.exe", , 2)
    Sleep(800)
    SendText(text)
}
ExitApp()
