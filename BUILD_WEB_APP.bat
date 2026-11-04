@echo off
setlocal
cd /d "%~dp0"
node "%~dp0build-web-app.js"
if errorlevel 1 (
  echo Failed to create the browser app download.
  exit /b 1
)
echo.
echo Browser app ZIP is ready in the downloads folder.
pause
