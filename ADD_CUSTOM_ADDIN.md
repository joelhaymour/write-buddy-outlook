# How to Add Your Custom Add-in (Not from Marketplace)

The Microsoft Marketplace you're seeing is for **published add-ins**. You need to add a **custom add-in from a file**.

## Steps to Add Your Custom Add-in:

### Step 1: Go Back to Outlook Desktop App

1. **Close the Marketplace website** (or switch back to Outlook app)

2. **In Outlook desktop app**, click **"File"** (top-left menu)

3. Look for one of these options:
   - **"Get Add-ins"** → Then click **"My Add-ins"** tab
   - **"Manage Add-ins"** → Then click **"My Add-ins"** tab
   - **"Options"** → **"Add-ins"** → Change dropdown to **"My Add-ins"** → Click **"Go"**

### Step 2: Add Custom Add-in

1. In the **"My Add-ins"** section, look for:
   - **"+ Add a Custom Add-in"** button (usually at the bottom)
   - Or **"Add from File"** option

2. Click **"+ Add a Custom Add-in"** → **"Add from File"**

3. Navigate to:
   ```
   /Users/joelhaymour/Downloads/Write Buddy Outlook/
   ```

4. Select **`manifest.xml`**

5. Click **"Add"** or **"Open"**

6. Wait for Outlook to validate and install it

### Step 3: Verify Installation

- You should see "Write Buddy" appear in your "My Add-ins" list
- The add-in should now be available in your Outlook ribbon

## If You Still Can't Find "My Add-ins"

**Alternative Method (Windows):**
1. File → Options → Add-ins
2. At the bottom, change dropdown from "COM Add-ins" to **"My Add-ins"**
3. Click **"Go"** button
4. Click **"+ Add a Custom Add-in"** → **"Add from File"**

**Alternative Method (Mac):**
1. Outlook menu → Preferences → Add-ins
2. Or: Tools → Add-ins → My Add-ins

## Important Notes

- **Marketplace** = Published add-ins (what you're seeing now)
- **My Add-ins** = Custom add-ins from files (what you need)
- Make sure the **server is running** (`./start-server.sh`) before using the add-in

