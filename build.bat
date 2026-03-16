@echo off
echo Сборка VoiceClaude.exe...
cd /d "%~dp0"
python -m PyInstaller --onefile --noconsole --name VoiceClaude --add-data "paste.ahk;." --collect-all faster_whisper --collect-all ctranslate2 main.py
echo.
echo Готово! Файл: dist\VoiceClaude.exe
pause
