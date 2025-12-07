# Check Browser Console for Errors

## The Issue
Outlook shows "installing" then fails. This means it's trying to validate the manifest but something is failing.

## What to Check

### 1. Open Browser Console in Outlook Web
1. In Chrome, while on Outlook Web
2. Press **F12** to open DevTools
3. Click the **Console** tab
4. Try installing the add-in again
5. **Look for any red error messages**
6. Take a screenshot or copy the error messages

### 2. Check Network Tab
1. In DevTools, click the **Network** tab
2. Try installing the add-in again
3. Look for any **failed requests** (they'll be red)
4. Click on failed requests to see the error details
5. Check which URL is failing

### 3. Common Issues to Look For
- **CORS errors**: "Access-Control-Allow-Origin" errors
- **404 errors**: File not found
- **SSL errors**: Certificate issues
- **JavaScript errors**: Errors in taskpane.js
- **Office.js errors**: Office.js not loading

## What We Need From You

Please share:
1. **Any error messages** from the Console tab
2. **Any failed network requests** from the Network tab
3. **The exact error message** Outlook shows (if any)

This will help us identify exactly what's failing!

