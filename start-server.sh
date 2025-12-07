#!/bin/bash

# Start the HTTPS server for Write Buddy Outlook Add-in

echo "🌐 Starting Write Buddy server..."
echo "Server will run on: https://localhost:3000"
echo "Press Ctrl+C to stop the server"
echo ""

# Check if certificates exist in default location
CERT_DIR="$HOME/.office-addin-dev-certs"
CERT_FILE="$CERT_DIR/localhost.crt"
KEY_FILE="$CERT_DIR/localhost.key"

# Check if certificates exist locally first
if [ -f "cert.pem" ] && [ -f "key.pem" ]; then
    CERT_FILE="cert.pem"
    KEY_FILE="key.pem"
    echo "✅ Using local certificates"
elif [ -f "$CERT_FILE" ] && [ -f "$KEY_FILE" ]; then
    echo "✅ Using certificates from: $CERT_DIR"
elif [ -f "$CERT_FILE" ] && [ -f "$CERT_DIR/localhost.key" ]; then
    KEY_FILE="$CERT_DIR/localhost.key"
    echo "✅ Using certificates from: $CERT_DIR"
else
    echo "⚠️  SSL certificates not found. Generating them..."
    npx office-addin-dev-certs install --machine
    if [ ! -f "$CERT_FILE" ] || [ ! -f "$KEY_FILE" ]; then
        echo "❌ Failed to generate certificates"
        echo "   Looking for: $CERT_FILE and $KEY_FILE"
        exit 1
    fi
fi

# Start the server
echo ""
npx http-server -S -C "$CERT_FILE" -K "$KEY_FILE" -p 3000 -c-1

