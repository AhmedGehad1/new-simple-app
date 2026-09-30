@echo off
rem Builds dist\CertificatePDFBuilder.exe (needs Python 3.9+ from python.org).
setlocal
cd /d "%~dp0"

where py >nul 2>nul && (set "PY=py -3") || (set "PY=python")
if not exist .venv\Scripts\python.exe (
    echo Creating a virtual environment...
    %PY% -m venv .venv || goto :error
)
.venv\Scripts\python.exe -m pip install --upgrade pip || goto :error
.venv\Scripts\python.exe -m pip install -r requirements.txt pyinstaller || goto :error
.venv\Scripts\python.exe tools\build_exe.py || goto :error

echo.
echo Done! The program is dist\CertificatePDFBuilder.exe
pause
exit /b 0

:error
echo.
echo The build failed - see the messages above.
pause
exit /b 1
