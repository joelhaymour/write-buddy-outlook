#!/bin/bash

# Update manifest.xml with ngrok URL

if [ -z "$1" ]; then
    echo "Usage: ./update-manifest-ngrok.sh <ngrok-url>"
    echo "Example: ./update-manifest-ngrok.sh https://abc123.ngrok.io"
    exit 1
fi

NGROK_URL="$1"

# Remove trailing slash if present
NGROK_URL="${NGROK_URL%/}"

echo "🔄 Updating manifest.xml with ngrok URL: $NGROK_URL"
echo ""

# Backup original
cp manifest.xml manifest.xml.backup

# Replace localhost:3000 with ngrok URL
sed -i '' "s|https://localhost:3000|$NGROK_URL|g" manifest.xml

echo "✅ Updated manifest.xml"
echo "📋 Backup saved as: manifest.xml.backup"
echo ""
echo "Next steps:"
echo "1. Transfer updated manifest.xml to Windows"
echo "2. Install add-in using the updated manifest.xml"
echo ""

