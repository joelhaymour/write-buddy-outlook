@echo off
REM Setup script for Write Buddy Outlook Add-in on Windows

echo Setting up Write Buddy Outlook Add-in...
echo.

REM Check if Node.js is installed
where node >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo ERROR: Node.js is not installed.
    echo Please install Node.js from https://nodejs.org/
    pause
    exit /b 1
)

echo Node.js found: 
node --version
echo npm found:
npm --version
echo.

REM Install dependencies
echo Installing dependencies...
call npm install
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo WARNING: npm install had issues. Trying to continue...
)

echo.
echo Generating SSL certificate...
call npx office-addin-dev-certs install --machine

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo WARNING: Certificate generation had issues.
    echo You may need to run: npx office-addin-dev-certs install --machine
)

echo.
echo Setup complete!
echo.
echo Next steps:
echo 1. Start the server: start-server.bat
echo 2. Load manifest.xml in Outlook (File -^> Get Add-ins -^> Add from File)
echo.
pause

