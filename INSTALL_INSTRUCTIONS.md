# How to Install Write Buddy in Outlook

## Step 1: Make Sure the Server is Running

1. Open Terminal (or the terminal window where you ran `./start-server.sh`)
2. Navigate to the folder:
   ```bash
   cd "/Users/joelhaymour/Downloads/Write Buddy Outlook"
   ```
3. Start the server:
   ```bash
   ./start-server.sh
   ```
4. **Keep this terminal window open!** The server must be running for the add-in to work.

You should see:
```
🌐 Starting Write Buddy server...
Server will run on: https://localhost:3000
```

## Step 2: Load the Add-in in Outlook

1. **Open Outlook** (the desktop app, not Outlook on the web)

2. **Go to File menu** (top-left corner)

3. **Click "Get Add-ins"** (or "Manage Add-ins" in some versions)
   - This opens the Office Add-ins dialog

4. **Click the "My Add-ins" tab** at the top

5. **Click the "+ Add a Custom Add-in" button** (usually at the bottom)

6. **Select "Add from File"** from the dropdown

7. **Navigate to and select the manifest.xml file:**
   - Go to: `/Users/joelhaymour/Downloads/Write Buddy Outlook/`
   - Select `manifest.xml`
   - Click "Add" or "Open"

8. **Wait for installation** - Outlook will validate and install the add-in

9. **You should see "Write Buddy" appear in your add-ins list**

## Step 3: Use the Add-in

1. **Compose a new email** or **reply to an email**

2. **Look for "Write Buddy" button** in the ribbon (usually in the "Home" tab)

3. **Click the "Write Buddy" button** to open the task pane

4. **Select text** in your email body

5. **Click "Get Selected Text"** in the task pane

6. **Choose an option** (Professional, Fact Check, Persuasive, or Translate)

7. **Click "Insert"** to replace the original text

## Troubleshooting

### Add-in Not Appearing?
- Make sure the server is running (`./start-server.sh`)
- Check that you're using Outlook desktop app (not web version)
- Try restarting Outlook after installation

### Can't Find "Get Add-ins"?
- Some Outlook versions have it under **File → Manage Add-ins**
- Or look for **File → Options → Add-ins → Go** (then click "My Add-ins" tab)

### Add-in Shows Error?
- Make sure the server is running on `https://localhost:3000`
- Check that the manifest.xml URLs point to `https://localhost:3000`
- Open browser DevTools (F12) in the task pane to see error messages

### "This add-in cannot be started"?
- Verify the server is running
- Check that `https://localhost:3000` is accessible in your browser
- Make sure SSL certificates are properly installed

## Next Steps

After installation:
1. **Configure your API key** - Click "Configure API Key" in the task pane
2. **Enable/disable features** - Use the checkboxes in Settings
3. **Start rewriting!** - Select text and use the options

