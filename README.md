# Write Buddy - Outlook Add-in

An Outlook add-in for Windows that rewrites selected text using the OpenAI API. Same functionality as the Chrome extension, but designed for Outlook desktop app.

## Features

- ✍️ **Professional**: Improves wording and clarity while preserving all facts
- ✅ **Fact Check**: Verifies facts, corrects errors, and improves writing
- 💼 **Persuasive**: Makes text more sales-oriented and compelling
- 🌐 **Translate**: Translates text to 24+ languages
- 🔒 **Secure**: API key stored in Outlook roaming settings
- ⚙️ **Customizable**: Enable/disable features via toggles

## Installation

### Prerequisites

- Outlook for Windows (desktop app)
- Node.js (for local development server)
- SSL certificate (for HTTPS - required by Outlook)

### Step 1: Set Up Local Server

Outlook add-ins require HTTPS. You'll need to:

1. **Install Node.js** (if not already installed)
   - Download from https://nodejs.org/
   - Install it

2. **Set up a local HTTPS server**
   - You can use tools like `http-server` with SSL
   - Or use `office-addin-dev-certs` for development certificates

### Step 2: Update Manifest

1. **Open `manifest.xml`**
2. **Replace `YOUR-ADD-IN-ID-HERE`** with a unique GUID
   - Generate one at https://www.guidgenerator.com/
3. **Update all `https://localhost:3000` URLs** to match your server URL

### Step 3: Sideload the Add-in

1. **Open Outlook**
2. **Go to File → Manage Add-ins** (or **Get Add-ins**)
3. **Click "My Add-ins"** → **"Add a Custom Add-in"** → **"Add from File"**
4. **Select the `manifest.xml` file**
5. The add-in should now appear in your Outlook ribbon

### Step 4: Configure API Key

1. **Open the Write Buddy task pane** (click the button in Outlook ribbon)
2. **Click "Configure API Key"**
3. **Enter your OpenAI API key** (from https://platform.openai.com/api-keys)
4. **Click OK**

## Usage

1. **Open an email** (compose or reply)
2. **Click "Write Buddy"** button in the ribbon
3. **Select text** in the email body
4. **Click "Get Selected Text"** in the task pane
5. **Choose an option**: Professional, Fact Check, Persuasive, or Translate
6. **Review** the rewritten text side-by-side
7. **Click "Insert"** to replace the original text

## File Structure

```
Write Buddy Outlook/
├── manifest.xml          # Outlook add-in manifest
├── taskpane.html         # Main UI
├── taskpane.js           # Main logic
├── taskpane.css          # Styling
└── README.md             # This file
```

## Development Setup

### Using Office Add-in Dev Tools

1. **Install Office Add-in Dev Tools:**
   ```bash
   npm install -g office-addin-dev-certs
   ```

2. **Generate SSL certificate:**
   ```bash
   office-addin-dev-certs install
   ```

3. **Start a local server:**
   ```bash
   npx http-server -S -C cert.pem -K key.pem -p 3000
   ```

4. **Update manifest.xml** with your local URL

### Testing

1. Sideload the add-in in Outlook
2. Open the task pane
3. Test all features
4. Check the browser console (F12) for errors

## Publishing to AppSource

To publish to Microsoft AppSource:

1. **Create a Microsoft Partner Center account**
2. **Package your add-in** (create a ZIP file)
3. **Submit for certification**
4. **Wait for approval** (usually 1-2 weeks)

## Differences from Chrome Extension

- Uses **Office.js API** instead of Chrome APIs
- **Task pane** instead of floating menu
- **Roaming settings** instead of Chrome storage
- Requires **HTTPS server** (even for local development)
- **Manifest.xml** instead of manifest.json

## Troubleshooting

### Add-in Not Loading

- Check that your server is running on HTTPS
- Verify the manifest.xml URLs are correct
- Check Outlook's add-in settings

### API Key Not Saving

- Ensure you have proper permissions
- Check browser console for errors
- Try restarting Outlook

### Text Selection Not Working

- Make sure you're in compose/edit mode
- Try selecting text first, then clicking "Get Selected Text"
- Check that Office.js is loaded properly

## Support

For issues or questions, check:
- Office.js documentation: https://docs.microsoft.com/en-us/office/dev/add-ins/
- Outlook add-in development: https://docs.microsoft.com/en-us/outlook/add-ins/

