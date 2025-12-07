# Sideloading Rejected by Exchange - Solutions

## The Problem
Your organization's Exchange server has **disabled sideloading** of custom add-ins. This is an IT policy restriction, not a problem with the manifest.

Error: `"Sideloading rejected by Exchange"`

## Solutions

### Option 1: Contact IT Administrator (Recommended for Work Accounts)
Ask your IT administrator to enable sideloading for your account:

**What to ask:**
- "Can you enable Outlook add-in sideloading for my account?"
- "I need to install a custom Outlook add-in for development/testing"
- They may need to run: `Set-OrganizationConfig -SideloadingEnabled $true` in Exchange PowerShell

**Note:** Many organizations disable this for security reasons, so they may say no.

### Option 2: Use Personal Microsoft Account
If you have a personal Microsoft account (outlook.com, hotmail.com, live.com):

1. Go to **outlook.live.com** or **outlook.office.com**
2. Sign in with your **personal account** (not work account)
3. Try installing the add-in there
4. Personal accounts typically allow sideloading

### Option 3: Use Outlook Desktop with Personal Account
1. Add your personal Microsoft account to Outlook Desktop
2. Switch to that account
3. Try installing the add-in
4. Personal accounts usually allow sideloading

### Option 4: Publish to Office Store (Long-term Solution)
If you want to distribute this add-in:
1. Publish it through the Office Store
2. Requires Microsoft approval process
3. Takes time but works for all users
4. See: https://learn.microsoft.com/en-us/office/dev/store/submit-to-the-office-store

### Option 5: Use Developer Account
If you have access to a Microsoft 365 Developer account:
- These typically allow sideloading
- Good for testing and development

## Quick Test
To verify if it's the account:
1. Try with a **personal Microsoft account** (outlook.com)
2. If it works there, it confirms the work account is blocking it
3. If it still fails, there may be another issue

## Next Steps
1. **Try with a personal account first** (easiest test)
2. If that works, you know it's the work account restriction
3. Then decide: contact IT, use personal account, or publish to store

