# Transferring Write Buddy to Windows

## Quick Transfer Guide

### Method 1: Cloud Storage (Easiest)

1. **Upload the folder to OneDrive/Google Drive/Dropbox:**
   - Upload: `/Users/joelhaymour/Downloads/Write Buddy Outlook/`
   - Download on Windows computer

2. **On Windows:**
   - Extract/download the folder
   - Follow the Windows setup steps

### Method 2: USB Drive

1. **Copy folder to USB:**
   - Copy: `/Users/joelhaymour/Downloads/Write Buddy Outlook/`
   - Paste to USB drive

2. **On Windows:**
   - Copy from USB to a folder (e.g., `C:\Write Buddy Outlook\`)
   - Follow the Windows setup steps

### Method 3: Network Share (If on Same Network)

1. **Share the folder from Mac**
2. **Access from Windows**
3. **Copy to Windows**

## Files to Transfer

Make sure you transfer the entire folder with all these files:
- `manifest.xml`
- `taskpane.html`
- `taskpane.js`
- `taskpane.css`
- `commands.html`
- `package.json`
- `start-server.bat` (Windows script)
- `setup-windows.bat` (Windows setup script)
- All other files

## On Windows - Quick Setup

1. **Open the folder in File Explorer**

2. **Right-click `setup-windows.bat`** → **Run as Administrator**

3. **Wait for setup to complete**

4. **Double-click `start-server.bat`** to start the server

5. **Install in Outlook:**
   - File → Get Add-ins → My Add-ins
   - + Add a Custom Add-in → Add from File
   - Select `manifest.xml`

## That's It!

The server will run on `https://localhost:3000` and the add-in will work!

