@echo off
REM Build the single-file Windows .exe (native desktop app).
REM
REM Packages app\report_generator_app.py - the Tkinter desktop UI. This is
REM a real Windows window: no browser, no HTML, and no network port is
REM opened at all. Each user runs their own copy; there is no host PC.
REM
REM (The browser/LAN version app\web_app.py still runs from source with
REM `python app\web_app.py`, but is no longer what gets packaged.)
REM
REM Optional code signing - set these before running to sign the build:
REM   set SIGN_CERT=C:\path\to\cert.pfx
REM   set CERT_PASSWORD=...

setlocal
cd /d "%~dp0"

if not exist .venv (
  py -m venv .venv
)

call .venv\Scripts\activate
python -m pip install --upgrade pip setuptools
python -m pip install -r requirements.txt pyinstaller pip-audit bandit
if errorlevel 1 (
  echo.
  echo ERROR: Dependency install failed.
  exit /b 1
)

echo.
echo ============================================================
echo  SECURITY GATE - the build stops if either check fails
echo ============================================================
echo.

echo [1/2] Scanning dependencies for known vulnerabilities...
pip-audit --progress-spinner off
if errorlevel 1 (
  echo.
  echo BUILD STOPPED: vulnerable dependencies found.
  echo Bump the affected pins in requirements.txt to the listed fix
  echo version, re-test, and run this script again. Do not ship this.
  exit /b 1
)
echo     OK - no known vulnerabilities.
echo.

echo [2/2] Static security scan of the code being shipped...
REM Medium severity and above only; the remaining Low findings are
REM defensive try/except blocks, reviewed and accepted.
REM web_app.py is excluded because it is not part of this .exe - it is
REM the browser version, and its one finding (binding 0.0.0.0) does not
REM exist in the desktop app being packaged here.
bandit -r app tools -ll --exclude app/web_app.py
if errorlevel 1 (
  echo.
  echo BUILD STOPPED: medium or high severity code finding.
  exit /b 1
)
echo     OK - no medium or high severity findings.
echo.

if not exist dist mkdir dist

echo Building single-file executable...
pyinstaller --noconfirm --clean --windowed --onefile ^
  --name SKF_Report_Generator ^
  --version-file version_info.txt ^
  --add-data "assets\Project Specification - Template.docx;assets" ^
  --add-data "assets\Project Specification - Decision Rule Source.docx;assets" ^
  --add-data "assets\Final Test Report - Template.docx;assets" ^
  --collect-all tkinterdnd2 ^
  --collect-all tkcalendar ^
  --collect-all babel ^
  --exclude-module flask ^
  --exclude-module waitress ^
  app\report_generator_app.py
if errorlevel 1 (
  echo.
  echo ERROR: PyInstaller build failed.
  exit /b 1
)

echo.
if defined SIGN_CERT (
  echo Signing the executable...
  signtool sign /f "%SIGN_CERT%" /p "%CERT_PASSWORD%" /fd sha256 ^
    /tr http://timestamp.digicert.com /td sha256 ^
    dist\SKF_Report_Generator.exe
  if errorlevel 1 (
    echo.
    echo ERROR: Signing failed. The .exe exists but is UNSIGNED.
    exit /b 1
  )
  echo     Signed.
) else (
  echo ***********************************************************
  echo  WARNING: SIGN_CERT not set - this build is UNSIGNED.
  echo.
  echo  Windows SmartScreen will show an "unknown publisher"
  echo  warning, and antivirus may quarantine it outright.
  echo  Ask IT for a code-signing certificate, then re-run with:
  echo     set SIGN_CERT=C:\path\to\cert.pfx
  echo     set CERT_PASSWORD=your-password
  echo ***********************************************************
)

echo.
echo Computing SHA256 checksum for integrity verification...
certutil -hashfile dist\SKF_Report_Generator.exe SHA256 > dist\SKF_Report_Generator.exe.sha256.txt
type dist\SKF_Report_Generator.exe.sha256.txt

echo.
echo Build completed.
echo Executable: dist\SKF_Report_Generator.exe
echo Checksum:   dist\SKF_Report_Generator.exe.sha256.txt
echo.
echo Share the checksum alongside the .exe (e.g. in Teams/email) so
echo recipients can verify the file matches what was built here:
echo   certutil -hashfile SKF_Report_Generator.exe SHA256
endlocal
