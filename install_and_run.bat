@echo off
chcp 65001 > nul
cd /d "%~dp0"

if not exist ".venv" (
    py -3 -m venv .venv
)

call .venv\Scripts\activate
python -m pip install --upgrade pip
pip install -r requirements.txt
python -m playwright install chromium
python main.py

pause
