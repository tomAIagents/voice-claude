@echo off
setlocal EnableDelayedExpansion
chcp 65001 >nul
echo ================================================
echo   Ustanovka VoiceClaude
echo ================================================
echo.

:: ── 1. Python ──────────────────────────────────────────────────────────────
set PYTHON_EXE=

:: Ищем python.exe, пропуская заглушку WindowsApps
for /f "delims=" %%i in ('where python 2^>nul') do (
    echo %%i | findstr /i "WindowsApps" >nul
    if errorlevel 1 (
        if "!PYTHON_EXE!"=="" set PYTHON_EXE=%%i
    )
)

:: Проверяем стандартные пути установки
if "!PYTHON_EXE!"=="" if exist "%LOCALAPPDATA%\Programs\Python\Python313\python.exe" set PYTHON_EXE=%LOCALAPPDATA%\Programs\Python\Python313\python.exe
if "!PYTHON_EXE!"=="" if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" set PYTHON_EXE=%LOCALAPPDATA%\Programs\Python\Python312\python.exe
if "!PYTHON_EXE!"=="" if exist "%LOCALAPPDATA%\Programs\Python\Python311\python.exe" set PYTHON_EXE=%LOCALAPPDATA%\Programs\Python\Python311\python.exe
if "!PYTHON_EXE!"=="" if exist "%LOCALAPPDATA%\Programs\Python\Python310\python.exe" set PYTHON_EXE=%LOCALAPPDATA%\Programs\Python\Python310\python.exe
if "!PYTHON_EXE!"=="" if exist "C:\Program Files\Python313\python.exe" set PYTHON_EXE=C:\Program Files\Python313\python.exe
if "!PYTHON_EXE!"=="" if exist "C:\Program Files\Python312\python.exe" set PYTHON_EXE=C:\Program Files\Python312\python.exe
if "!PYTHON_EXE!"=="" if exist "C:\Program Files\Python311\python.exe" set PYTHON_EXE=C:\Program Files\Python311\python.exe

if "!PYTHON_EXE!"=="" (
    echo [1/4] Python ne nayden - ustanavlivaem...
    winget install Python.Python.3.11 --source winget -e --accept-source-agreements --accept-package-agreements
    :: После установки winget обновляем PATH
    for /f "skip=2 tokens=3*" %%a in ('reg query "HKCU\Environment" /v PATH 2^>nul') do set USER_PATH=%%a %%b
    for /f "skip=2 tokens=3*" %%a in ('reg query "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v PATH 2^>nul') do set SYS_PATH=%%a %%b
    set PATH=!SYS_PATH!;!USER_PATH!
    :: Ищем снова после установки
    for /f "delims=" %%i in ('where python 2^>nul') do (
        echo %%i | findstr /i "WindowsApps" >nul
        if errorlevel 1 (
            if "!PYTHON_EXE!"=="" set PYTHON_EXE=%%i
        )
    )
    if "!PYTHON_EXE!"=="" if exist "%LOCALAPPDATA%\Programs\Python\Python311\python.exe" set PYTHON_EXE=%LOCALAPPDATA%\Programs\Python\Python311\python.exe
) else (
    echo [1/4] Python - uzhe ustanovlen
)

if "!PYTHON_EXE!"=="" (
    echo.
    echo OSHIBKA: Python ne nayden.
    echo Perezapusti install.bat posle perekhoda po ssylke:
    echo https://www.python.org/downloads/
    pause
    exit /b 1
)
echo     Ispolzuetsya: !PYTHON_EXE!

:: ── 2. AutoHotkey ──────────────────────────────────────────────────────────
echo [2/4] Ustanovka AutoHotkey v2...
winget install AutoHotkey.AutoHotkey --source winget -e --accept-source-agreements --accept-package-agreements 2>nul
if errorlevel 1 echo     AutoHotkey uzhe ustanovlen.

:: ── 3. Python-библиотеки ───────────────────────────────────────────────────
echo [3/4] Ustanovka bibliotek...
"!PYTHON_EXE!" -m pip install --upgrade pip --quiet
"!PYTHON_EXE!" -m pip install faster-whisper keyboard pyaudio
if errorlevel 1 (
    echo     pyaudio ne ustanovilsya napryamuyu, probuyem cherez pipwin...
    "!PYTHON_EXE!" -m pip install pipwin --quiet
    "!PYTHON_EXE!" -m pipwin install pyaudio
)

:: ── 4. Ярлык на рабочем столе ─────────────────────────────────────────────
echo [4/4] Sozdanie yarlyika na rabochem stole...
set "SCRIPT_DIR=%~dp0"
set "DESKTOP=%USERPROFILE%\Desktop"
set "VBS_PATH=%SCRIPT_DIR%VoiceClaude.vbs"

powershell -Command "$ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut('%DESKTOP%\VoiceClaude.lnk'); $s.TargetPath = '%VBS_PATH%'; $s.Description = 'VoiceClaude'; $s.Save()"

echo.
echo ================================================
echo   Gotovo! Yarlyik sozdan na rabochem stole.
echo.
echo   VAZHNO: Pri pervom zapuske programma skachaet
echo   model (~1.5 GB). Eto займет neskolko minut.
echo   Okno pokazhet "Zagruzka modeli..." - eto normalno.
echo ================================================
pause
