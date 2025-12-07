# Setting Up on Windows Computer

## Option 1: Transfer Files to Windows (RECOMMENDED - Easiest)

### Step 1: Transfer Files
1. **Copy the entire folder** to your Windows computer:
   - Use USB drive, cloud storage (OneDrive, Google Drive), or network share
   - Folder: `/Users/joelhaymour/Downloads/Write Buddy Outlook/`

### Step 2: On Windows Computer
1. **Install Node.js** (if not installed):
   - Download from https://nodejs.org/
   - Install it

2. **Open PowerShell or Command Prompt** in the folder:
   ```powershell
   cd "C:\path\to\Write Buddy Outlook"
   ```

3. **Install dependencies:**
   ```powershell
   npm install
   ```

4. **Generate SSL certificate:**
   ```powershell
   npx office-addin-dev-certs install --machine
   ```

5. **Start the server:**
   ```powershell
   npx http-server -S -C "%USERPROFILE%\.office-addin-dev-certs\localhost.crt" -K "%USERPROFILE%\.office-addin-dev-certs\localhost.key" -p 3000 -c-1
   ```

   Or use the start-server.sh script (if you have Git Bash or WSL):
   ```bash
   ./start-server.sh
   ```

6. **Keep the server running**

### Step 3: Install Add-in in Outlook (Windows)
1. **Open Outlook** (desktop app)
2. **File** → **Get Add-ins** (or **Manage Add-ins**)
3. **My Add-ins** tab
4. **+ Add a Custom Add-in** → **Add from File**
5. Select `manifest.xml`
6. **Add**

## Option 2: Use ngrok (Keep Server on Mac)

### Step 1: On Mac - Install ngrok
```bash
# Install ngrok (if not installed)
brew install ngrok
# Or download from https://ngrok.com/download
```

### Step 2: On Mac - Start Server
```bash
cd "/Users/joelhaymour/Downloads/Write Buddy Outlook"
./start-server.sh
# Keep this running in one terminal
```

### Step 3: On Mac - Start ngrok
```bash
# In a new terminal window
ngrok http 3000
```

This will give you a URL like: `https://abc123.ngrok.io`

### Step 4: Update manifest.xml
Update all URLs in manifest.xml from `https://localhost:3000` to your ngrok URL:
- `https://abc123.ngrok.io` (use your actual ngrok URL)

### Step 5: Transfer Updated manifest.xml to Windows
1. Copy the updated `manifest.xml` to Windows
2. Install the add-in using the updated manifest

### Step 6: Keep Both Running
- Keep the server running on Mac
- Keep ngrok running on Mac
- Use Outlook on Windows

## Recommendation

**Use Option 1** (transfer to Windows) - it's simpler and doesn't require:
- Keeping Mac running
- ngrok setup
- URL updates

## Quick Windows Setup Script

I can create a Windows batch file to make it easier. Would you like me to create that?

