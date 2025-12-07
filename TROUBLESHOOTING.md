# Troubleshooting Outlook Add-in Installation

## Common Installation Failures

### 1. "Installation Failed" Error

**Possible Causes:**
- Manifest XML validation errors
- URLs not accessible from Windows
- SSL certificate issues
- Missing required files

### 2. Check These Things:

#### On Your Mac (Server Side):
1. **Verify server is running:**
   ```bash
   ps aux | grep http-server
   ```

2. **Verify ngrok is running:**
   ```bash
   curl http://localhost:4040/api/tunnels
   ```

3. **Test URLs are accessible:**
   - Open browser on Windows and visit: `https://YOUR_NGROK_URL.ngrok.app/taskpane.html`
   - Should see the Write Buddy interface

#### On Windows (Client Side):
1. **Check if you can access the ngrok URL:**
   - Open Internet Explorer or Edge
   - Visit: `https://YOUR_NGROK_URL.ngrok.app/taskpane.html`
   - If you see an SSL warning, click "Advanced" → "Continue to site"

2. **Check Outlook version:**
   - Outlook 2016 or later required
   - Office 365 or Outlook 2019/2021 recommended

3. **Check manifest file:**
   - Make sure you downloaded the latest version from GitHub
   - The manifest.xml should have the current ngrok URL

### 3. Common Fixes:

#### Fix 1: Update Manifest URL
If ngrok restarted and you have a new URL:
1. On Mac, run: `./update-manifest-ngrok.sh <new-url>`
2. Download updated manifest.xml from GitHub
3. Try installing again

#### Fix 2: Check File Accessibility
Test each URL in the manifest:
- `https://YOUR_NGROK_URL.ngrok.app/taskpane.html` - Should load
- `https://YOUR_NGROK_URL.ngrok.app/assets/icon-32.png` - Should download
- `https://YOUR_NGROK_URL.ngrok.app/commands.html` - Should load

#### Fix 3: Clear Outlook Cache
1. Close Outlook
2. Delete: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`
3. Restart Outlook
4. Try installing again

#### Fix 4: Check Windows Firewall
- Make sure Windows Firewall isn't blocking the connection
- Try temporarily disabling firewall to test

### 4. Get More Error Details:

1. **In Outlook:**
   - File → Options → Trust Center → Trust Center Settings → Privacy Options
   - Enable "Allow the use of connected experiences in Office"
   - Check "Enable logging"

2. **Check Event Viewer:**
   - Windows Key + X → Event Viewer
   - Look for Office/Outlook errors

3. **Check Outlook Logs:**
   - `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\Logs\`

### 5. Alternative: Use Outlook Web

If desktop Outlook doesn't work, try Outlook Web:
1. Go to outlook.office.com
2. Settings → View all Outlook settings → Mail → Add-ins
3. Upload manifest.xml

### 6. Still Not Working?

1. **Verify manifest XML is valid:**
   - Open manifest.xml in a text editor
   - Check for any obvious errors
   - Make sure all URLs use HTTPS

2. **Test with a simpler manifest:**
   - Try removing optional elements (icons, support URL)
   - Keep only required elements

3. **Check ngrok status:**
   - Visit http://localhost:4040 on Mac
   - Check if tunnel is active
   - Verify requests are coming through

