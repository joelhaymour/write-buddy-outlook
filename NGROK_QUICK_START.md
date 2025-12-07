# 🚀 Quick Start: Using ngrok

## ✅ Current Status

- **Server**: Running on `https://localhost:3000`
- **ngrok URL**: `https://40d12b37b455.ngrok.app`
- **Manifest**: Updated with ngrok URL

## 📋 Next Steps

### 1. Transfer manifest.xml to Windows

Transfer the updated `manifest.xml` file to your Windows computer. You can:
- Use GitHub (push to repo, pull on Windows)
- Use a USB drive
- Email it to yourself
- Use cloud storage (OneDrive, Dropbox, etc.)

### 2. Install Add-in on Windows

1. Open Outlook on Windows
2. Go to **File** → **Get Add-ins** → **My Add-ins** → **Custom Add-ins** → **Add a Custom Add-in** → **Add from file**
3. Select the `manifest.xml` file
4. Click **Add**

### 3. Use the Add-in

1. Open a new email in Outlook
2. Click the **Write Buddy** button in the ribbon
3. Select text in your email
4. Choose an option (Professional, Fact Check, Persuasive, or Translate)
5. Click **Insert** to replace the text

## ⚠️ Important Notes

- **Keep your Mac running**: The server must stay running on your Mac for the add-in to work on Windows
- **ngrok URL changes**: If you restart ngrok, you'll get a new URL. You'll need to:
  1. Run `./update-manifest-ngrok.sh <new-url>` on Mac
  2. Transfer the updated `manifest.xml` to Windows again
  3. Reinstall the add-in in Outlook

## 🔄 If You Need to Restart

If you need to restart the server or ngrok:

1. **Stop current processes**:
   ```bash
   pkill -f "http-server"
   pkill -f "ngrok"
   ```

2. **Start server**:
   ```bash
   cd "/Users/joelhaymour/Downloads/Write Buddy Outlook"
   ./start-server.sh &
   ```

3. **Start ngrok** (in a new terminal):
   ```bash
   ngrok http 3000
   ```

4. **Get new ngrok URL**:
   - Visit http://localhost:4040 in your browser
   - Or run: `curl -s http://localhost:4040/api/tunnels | grep -o '"public_url":"[^"]*"' | head -1 | cut -d'"' -f4`

5. **Update manifest**:
   ```bash
   ./update-manifest-ngrok.sh <new-ngrok-url>
   ```

6. **Transfer updated manifest.xml to Windows and reinstall**

## 🐛 Troubleshooting

- **Add-in doesn't load**: Make sure the server is running on Mac and ngrok is active
- **Connection errors**: Check that the ngrok URL in manifest.xml matches your current ngrok URL
- **Can't access taskpane**: Verify the server is accessible at the ngrok URL in a browser

