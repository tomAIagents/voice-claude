@echo off
chcp 65001 >nul
echo ================================================
echo   Установка VoiceClaude
echo ================================================
echo.

:: Проверка Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [1/4] Установка Python...
    winget install Python.Python.3.11 --source winget -e
) else (
    echo [1/4] Python — уже установлен
)

:: Установка AutoHotkey
echo [2/4] Установка AutoHotkey...
winget install AutoHotkey.AutoHotkey --source winget -e

:: Установка Python-библиотек
echo [3/4] Установка библиотек...
pip install faster-whisper pyaudio keyboard pyperclip

:: Создание ярлыка на рабочем столе
echo [4/4] Создание ярлыка на рабочем столе...
set SCRIPT_DIR=%~dp0
set DESKTOP=%USERPROFILE%\Desktop
set VBS_PATH=%SCRIPT_DIR%VoiceClaude.vbs

powershell -Command "$ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut('%DESKTOP%\VoiceClaude.lnk'); $s.TargetPath = '%VBS_PATH%'; $s.Description = 'Голосовой ввод для Claude Code'; $s.Save()"

echo.
echo ================================================
echo   Готово! Ярлык создан на рабочем столе.
echo   Запускай VoiceClaude с рабочего стола.
echo ================================================
pause
