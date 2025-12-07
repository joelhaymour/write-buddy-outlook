# Fix npm Cache Permission Issue

If you're getting npm cache permission errors, here's how to fix it:

## Quick Fix (Recommended)

Run this command in Terminal:

```bash
sudo chown -R $(whoami) ~/.npm
```

Then try the setup again:

```bash
cd "Write Buddy Outlook"
./setup.sh
```

## Alternative: Use npm with Different Cache

If you don't want to change ownership, you can use a different cache location:

```bash
npm install --cache ~/.npm-cache-alternative
```

## Manual Setup (If npm install still fails)

1. **Install packages globally:**
   ```bash
   npm install -g office-addin-dev-certs http-server
   ```

2. **Generate certificate:**
   ```bash
   office-addin-dev-certs install --machine
   ```

3. **Start server manually:**
   ```bash
   http-server -S -C cert.pem -K key.pem -p 3000 -c-1
   ```

## Why This Happens

The npm cache directory sometimes gets incorrect permissions, especially if npm was run with sudo at some point. The fix above restores proper ownership to your user account.

