# Outlook Add-in Setup Guide

## Quick Setup Steps

### 1. Install Dependencies

```bash
npm install
```

This installs:
- `office-addin-dev-certs` - For SSL certificates
- `http-server` - For local HTTPS server

### 2. Generate SSL Certificate

```bash
npm run cert
```

This creates SSL certificates needed for local development.

### 3. Start Local Server

```bash
npm start
```

This starts an HTTPS server on `https://localhost:3000`

### 4. Update Manifest

1. Open `manifest.xml`
2. Replace `12345678-1234-1234-1234-123456789abc` with a unique GUID
   - Generate one at: https://www.guidgenerator.com/
3. Make sure all URLs point to `https://localhost:3000` (or your server URL)

### 5. Sideload in Outlook

1. Open Outlook (Windows desktop app)
2. Go to **File** → **Get Add-ins** (or **Manage Add-ins**)
3. Click **My Add-ins** tab
4. Click **+ Add a Custom Add-in** → **Add from File**
5. Select `manifest.xml`
6. Click **Add**

### 6. Use the Add-in

1. Compose a new email or reply
2. Look for **"Write Buddy"** button in the ribbon
3. Click it to open the task pane
4. Select text in your email
5. Click **"Get Selected Text"**
6. Choose an option and click **"Insert"**

## Important Notes

- **HTTPS Required**: Outlook add-ins require HTTPS, even for local development
- **GUID Required**: You must replace the placeholder GUID in manifest.xml with a unique one
- **Server Must Be Running**: The local server must be running when using the add-in
- **API Key**: Configure your OpenAI API key in the task pane settings

## Troubleshooting

### Add-in Not Appearing

- Make sure the server is running (`npm start`)
- Check that manifest.xml URLs are correct
- Verify SSL certificate is installed (`npm run cert`)

### Cannot Get Selected Text

- Make sure you're in compose/edit mode (not read mode)
- Try selecting text first, then clicking "Get Selected Text"
- Check browser console (F12) for errors

### API Errors

- Verify your API key is set correctly
- Check that the server can reach https://api.openai.com
- Look for CORS errors in the console

