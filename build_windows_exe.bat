@echo off
setlocal

REM Build Windows .exe using PyInstaller
if not exist .venv (
  py -m venv .venv
)

call .venv\Scripts\activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt pyinstaller

if not exist dist mkdir dist

pyinstaller --noconfirm --clean --windowed --onefile ^
  --name SKF_Report_Generator ^
  --add-data "assets\Project Specification - Template.docx;assets" ^
  app\report_generator_app.py

echo.
echo Build completed.
echo Executable: dist\SKF_Report_Generator.exe
endlocal
