# Using ngrok - No Node.js Needed on Windows!

## Quick Setup (3 Steps)

### Step 1: Start Server with ngrok on Mac

**On your Mac**, run:

```bash
cd "/Users/joelhaymour/Downloads/Write Buddy Outlook"
./start-with-ngrok.sh
```

This will:
- Start the server on localhost:3000
- Start ngrok tunnel
- Show you a public URL like: `https://abc123.ngrok.io`

**Keep this terminal window open!**

### Step 2: Update manifest.xml with ngrok URL

**In a new terminal window** (keep the first one running):

```bash
cd "/Users/joelhaymour/Downloads/Write Buddy Outlook"
./update-manifest-ngrok.sh https://YOUR-NGROK-URL.ngrok.io
```

Replace `YOUR-NGROK-URL` with the actual URL from Step 1.

This updates all URLs in manifest.xml to use the ngrok URL.

### Step 3: Install on Windows

1. **Transfer the updated `manifest.xml`** to Windows:
   - Email it to yourself
   - Or upload to OneDrive/Google Drive
   - Or push to GitHub and download

2. **On Windows**:
   - Open Outlook
   - File → Get Add-ins → My Add-ins
   - + Add a Custom Add-in → Add from File
   - Select the updated `manifest.xml`

3. **Done!** The add-in will work!

## Important Notes

- **Keep ngrok running on Mac** - if you close it, the URL stops working
- **ngrok free tier** gives you a random URL each time (changes when you restart)
- **ngrok paid tier** gives you a permanent URL
- **Server must be running** on Mac for it to work

## If ngrok URL Changes

If you restart ngrok and get a new URL:
1. Run `./update-manifest-ngrok.sh` again with the new URL
2. Transfer the updated manifest.xml to Windows
3. Reinstall the add-in (or it might auto-update)

## Alternative: Use Chrome Extension

If ngrok is too complicated:
- Use Outlook on the web (outlook.office.com)
- Use the Chrome extension you already have
- Right-click text → "Rewrite with Write Buddy"
- Works immediately, no setup needed!

