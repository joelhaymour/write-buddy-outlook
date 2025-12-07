@echo off
REM Start the HTTPS server for Write Buddy Outlook Add-in (Windows)

echo Starting Write Buddy server...
echo Server will run on: https://localhost:3000
echo Press Ctrl+C to stop the server
echo.

REM Check if certificates exist in default location
set CERT_DIR=%USERPROFILE%\.office-addin-dev-certs
set CERT_FILE=%CERT_DIR%\localhost.crt
set KEY_FILE=%CERT_DIR%\localhost.key

REM Check if certificates exist locally first
if exist "cert.pem" if exist "key.pem" (
    set CERT_FILE=cert.pem
    set KEY_FILE=key.pem
    echo Using local certificates
) else if exist "%CERT_FILE%" if exist "%KEY_FILE%" (
    echo Using certificates from: %CERT_DIR%
) else (
    echo SSL certificates not found. Generating them...
    call npx office-addin-dev-certs install --machine
    if not exist "%CERT_FILE%" (
        echo Failed to generate certificates
        pause
        exit /b 1
    )
)

REM Start the server
echo.
npx http-server -S -C "%CERT_FILE%" -K "%KEY_FILE%" -p 3000 -c-1

