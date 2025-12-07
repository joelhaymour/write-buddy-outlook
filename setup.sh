#!/bin/bash

# Write Buddy Outlook Add-in Setup Script

echo "🚀 Setting up Write Buddy Outlook Add-in..."
echo ""

# Check if Node.js is installed
if ! command -v node &> /dev/null; then
    echo "❌ Node.js is not installed. Please install it from https://nodejs.org/"
    exit 1
fi

echo "✅ Node.js found: $(node --version)"
echo "✅ npm found: $(npm --version)"
echo ""

# Fix npm cache permissions (if needed)
echo "🔧 Fixing npm cache permissions..."
npm cache clean --force 2>/dev/null || true

# Install dependencies
echo "📦 Installing dependencies..."
npm install

if [ $? -ne 0 ]; then
    echo ""
    echo "⚠️  npm install failed. Trying to fix cache permissions..."
    echo "   You may need to run: sudo chown -R $(whoami) ~/.npm"
    echo "   Or try: npm install --legacy-peer-deps"
    exit 1
fi

echo ""
echo "✅ Dependencies installed!"
echo ""

# Generate SSL certificate
echo "🔐 Generating SSL certificate..."
npx office-addin-dev-certs install --machine

if [ $? -ne 0 ]; then
    echo "⚠️  Certificate generation failed. You can try manually:"
    echo "   npx office-addin-dev-certs install"
    exit 1
fi

echo ""
echo "✅ SSL certificate generated!"
echo ""
echo "🎉 Setup complete!"
echo ""
echo "Next steps:"
echo "1. Start the server: npm start"
echo "2. Load manifest.xml in Outlook (File → Get Add-ins → Add from File)"
echo ""

