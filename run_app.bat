@echo off
chcp 65001 > nul
cd /d "%~dp0"

if not exist ".venv" (
    echo لم يتم العثور على البيئة الافتراضية. شغل install_and_run.bat أولاً.
    pause
    exit /b 1
)

call .venv\Scripts\activate
python main.py
