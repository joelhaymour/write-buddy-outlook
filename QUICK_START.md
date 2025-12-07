# Write Buddy Outlook Add-in - Quick Start

## What's Different from Chrome Extension?

The Outlook add-in works the same way but uses different technology:

| Feature | Chrome Extension | Outlook Add-in |
|---------|-----------------|----------------|
| **UI** | Right-click context menu → Modal | Task pane (sidebar) |
| **Text Selection** | Automatic on highlight | Click "Get Selected Text" button |
| **API** | Chrome APIs | Office.js API |
| **Storage** | Chrome sync storage | Outlook roaming settings |
| **Manifest** | manifest.json | manifest.xml |
| **Server** | Not needed | HTTPS server required |

## Quick Setup (5 Steps)

### 1. Install Node.js
Download from https://nodejs.org/ and install it.

### 2. Install Dependencies
Open terminal in the "Write Buddy Outlook" folder and run:
```bash
npm install
```

### 3. Generate SSL Certificate
```bash
npm run cert
```

### 4. Start Server
```bash
npm start
```
Keep this terminal window open! The server must be running.

### 5. Load in Outlook
1. Open Outlook (Windows desktop app)
2. **File** → **Get Add-ins** → **My Add-ins** tab
3. Click **+ Add a Custom Add-in** → **Add from File**
4. Select `manifest.xml` from this folder
5. Click **Add**

## Using the Add-in

1. **Compose or reply** to an email
2. **Click "Write Buddy"** button in the ribbon
3. **Select text** in your email body
4. **Click "Get Selected Text"** in the task pane
5. **Choose an option**: Professional, Fact Check, Persuasive, or Translate
6. **Review** the rewritten text
7. **Click "Insert"** to replace the original text

## Configure API Key

1. In the task pane, click **"Configure API Key"**
2. Enter your OpenAI API key (from https://platform.openai.com/api-keys)
3. Click **OK**

## Enable/Disable Features

Use the checkboxes in the Settings section to show/hide:
- Professional
- Fact Check
- Persuasive
- Translate

## Troubleshooting

**Add-in not appearing?**
- Make sure the server is running (`npm start`)
- Check that you're using Outlook desktop app (not web)

**Can't get selected text?**
- Make sure you're in compose/edit mode
- Select text first, then click "Get Selected Text"

**API errors?**
- Verify your API key is set correctly
- Check that the server can reach api.openai.com

## Next Steps

- See `SETUP.md` for detailed setup instructions
- See `README.md` for full documentation

