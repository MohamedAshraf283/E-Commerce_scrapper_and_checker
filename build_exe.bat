@echo off
chcp 65001 > nul
cd /d "%~dp0"

if not exist ".venv" (
    py -3 -m venv .venv
)
call .venv\Scripts\activate
python -m pip install --upgrade pip
pip install -r requirements.txt

pyinstaller --noconfirm --clean --windowed --name "SallaZidScraper" main.py

echo.
echo تم البناء داخل مجلد dist\SallaZidScraper
echo ملاحظة: Playwright يحتاج تثبيت المتصفح مرة واحدة على الجهاز:
echo python -m playwright install chromium
pause
