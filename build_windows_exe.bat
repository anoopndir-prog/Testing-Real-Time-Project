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
  --version-file version_info.txt ^
  --add-data "assets\Project Specification - Template.docx;assets" ^
  --add-data "assets\Project Specification - Decision Rule Source.docx;assets" ^
  app\report_generator_app.py

REM Optional: sign the build once you have a code-signing certificate.
REM signtool sign /f "path\to\cert.pfx" /p %CERT_PASSWORD% /fd sha256 /tr http://timestamp.digicert.com /td sha256 dist\SKF_Report_Generator.exe

echo.
echo Computing SHA256 checksum for integrity verification...
certutil -hashfile dist\SKF_Report_Generator.exe SHA256 > dist\SKF_Report_Generator.exe.sha256.txt
type dist\SKF_Report_Generator.exe.sha256.txt

echo.
echo Build completed.
echo Executable: dist\SKF_Report_Generator.exe
echo Checksum:   dist\SKF_Report_Generator.exe.sha256.txt
echo.
echo Share the checksum alongside the .exe (e.g. in Teams/email) so recipients
echo can verify the file they received matches what was built here, e.g.:
echo   certutil -hashfile SKF_Report_Generator.exe SHA256
endlocal
