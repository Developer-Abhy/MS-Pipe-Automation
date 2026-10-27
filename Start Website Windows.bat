@echo off
title MS Pipe Design Website
cd /d "%~dp0"
echo Checking Python...
python --version
if errorlevel 1 (
  echo Python was not found. Install Python 3.10 or later and select "Add Python to PATH".
  pause
  exit /b 1
)
echo.
echo Checking LibreOffice...
where soffice >nul 2>nul
if errorlevel 1 (
  if exist "C:\Program Files\LibreOffice\program\soffice.exe" (
    echo LibreOffice found.
  ) else if exist "C:\Program Files (x86)\LibreOffice\program\soffice.exe" (
    echo LibreOffice found.
  ) else (
    echo WARNING: LibreOffice was not found in PATH or its usual install folders.
    echo Install LibreOffice from https://www.libreoffice.org/download/download-libreoffice/
    echo Then run this file again.
    pause
    exit /b 1
  )
) else (
  echo LibreOffice found.
)
echo.
echo Installing required Python packages if needed...
python -m pip install -r requirements.txt
if errorlevel 1 (
  echo Package installation failed.
  pause
  exit /b 1
)
echo.
echo Starting website at http://127.0.0.1:5000
python app.py
pause
