# Outlook for Mac - Add-in Limitations

## Important: Outlook for Mac Limitations

**Outlook for Mac has limited support for custom add-ins** compared to Windows. Many versions don't support sideloading custom add-ins from manifest.xml files.

## Check Your Outlook Version

1. **Click "Outlook" in the menu bar** (top-left, next to Apple logo)
2. **Click "About Outlook"**
3. **Note the version number**

**Versions that MAY support custom add-ins:**
- Outlook for Mac 2016 (16.17+) - Limited support
- Outlook for Mac 2019 (16.43+) - Better support
- Outlook for Mac (Microsoft 365) - Best support

## Alternative Solutions

### Option 1: Use Outlook on the Web

1. **Go to**: https://outlook.office.com (or https://outlook.live.com)
2. **Sign in** with your account
3. **Click the gear icon** (⚙️) in top-right
4. **Click "View all Outlook settings"**
5. **Go to "Mail"** → **"Add-ins"**
6. **Click "My Add-ins"** tab
7. **Click "+ Add a Custom Add-in"** → **"Add from File"**
8. **Upload your manifest.xml**

**Note**: You'll still need the server running (`./start-server.sh`) and the manifest.xml URLs must point to your server.

### Option 2: Use Windows Outlook (If Available)

If you have access to a Windows computer:
1. Install the add-in there using the standard Windows method
2. It may sync to your Mac (depending on your setup)

### Option 3: Use the Chrome Extension Instead

Since you already have the Chrome extension working, you could:
- Use Outlook on the web in Chrome
- Use the Chrome extension there
- It works the same way!

### Option 4: Check if Your Version Supports It

Try this in Outlook for Mac:
1. **Compose a new email**
2. **Look at the ribbon/toolbar**
3. **See if there's an "Add-ins" tab or button**
4. If yes, click it and look for "My Add-ins"

## Recommendation

**For Mac users, I recommend using the Chrome extension with Outlook on the web** - it's easier and works reliably!

Would you like me to help you set that up instead?

