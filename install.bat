@echo off
chcp 65001 >nul
echo ================================================
echo   Установка VoiceClaude
echo ================================================
echo.

:: ── 1. Python ──────────────────────────────────────────────────────────────
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [1/4] Установка Python 3.11...
    winget install Python.Python.3.11 --source winget -e --accept-source-agreements --accept-package-agreements
) else (
    echo [1/4] Python — уже установлен
)

:: Ищем python.exe (PATH мог не обновиться после winget)
set PYTHON_EXE=
for /f "delims=" %%i in ('where python 2^>nul') do (
    if "!PYTHON_EXE!"=="" set PYTHON_EXE=%%i
)
if "%PYTHON_EXE%"=="" if exist "%LOCALAPPDATA%\Programs\Python\Python313\python.exe" set PYTHON_EXE=%LOCALAPPDATA%\Programs\Python\Python313\python.exe
if "%PYTHON_EXE%"=="" if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" set PYTHON_EXE=%LOCALAPPDATA%\Programs\Python\Python312\python.exe
if "%PYTHON_EXE%"=="" if exist "%LOCALAPPDATA%\Programs\Python\Python311\python.exe" set PYTHON_EXE=%LOCALAPPDATA%\Programs\Python\Python311\python.exe
if "%PYTHON_EXE%"=="" if exist "%LOCALAPPDATA%\Programs\Python\Python310\python.exe" set PYTHON_EXE=%LOCALAPPDATA%\Programs\Python\Python310\python.exe
if "%PYTHON_EXE%"=="" if exist "C:\Program Files\Python313\python.exe" set PYTHON_EXE=C:\Program Files\Python313\python.exe
if "%PYTHON_EXE%"=="" if exist "C:\Program Files\Python312\python.exe" set PYTHON_EXE=C:\Program Files\Python312\python.exe
if "%PYTHON_EXE%"=="" if exist "C:\Program Files\Python311\python.exe" set PYTHON_EXE=C:\Program Files\Python311\python.exe

if "%PYTHON_EXE%"=="" (
    echo.
    echo ОШИБКА: Python не найден. Перезапусти install.bat после установки Python.
    pause
    exit /b 1
)
echo     Используется: %PYTHON_EXE%

:: ── 2. AutoHotkey ──────────────────────────────────────────────────────────
echo [2/4] Установка AutoHotkey v2...
winget install AutoHotkey.AutoHotkey --source winget -e --accept-source-agreements --accept-package-agreements
if %errorlevel% neq 0 (
    echo     AutoHotkey уже установлен или не удалось установить — продолжаем.
)

:: ── 3. Python-библиотеки ───────────────────────────────────────────────────
echo [3/4] Установка библиотек...
"%PYTHON_EXE%" -m pip install --upgrade pip >nul
"%PYTHON_EXE%" -m pip install faster-whisper keyboard pyaudio
if %errorlevel% neq 0 (
    echo     Прямая установка pyaudio не удалась, пробуем через pipwin...
    "%PYTHON_EXE%" -m pip install pipwin
    "%PYTHON_EXE%" -m pipwin install pyaudio
)

:: ── 4. Ярлык на рабочем столе ─────────────────────────────────────────────
echo [4/4] Создание ярлыка на рабочем столе...
set SCRIPT_DIR=%~dp0
set DESKTOP=%USERPROFILE%\Desktop
set VBS_PATH=%SCRIPT_DIR%VoiceClaude.vbs

powershell -Command "$ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut('%DESKTOP%\VoiceClaude.lnk'); $s.TargetPath = '%VBS_PATH%'; $s.Description = 'Голосовой ввод для Claude Code'; $s.Save()"

echo.
echo ================================================
echo   Готово! Ярлык создан на рабочем столе.
echo.
echo   ВАЖНО: При первом запуске программа скачает
echo   языковую модель (~1.5 ГБ). Это займёт
echo   несколько минут — это нормально.
echo.
echo   Запускай VoiceClaude с рабочего стола.
echo ================================================
pause
