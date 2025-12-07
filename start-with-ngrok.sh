#!/bin/bash

# Start Write Buddy server with ngrok

echo "🚀 Starting Write Buddy with ngrok..."
echo ""

# Check if ngrok is installed
if ! command -v ngrok &> /dev/null; then
    echo "❌ ngrok is not installed."
    echo ""
    echo "Install it with:"
    echo "  brew install ngrok"
    echo ""
    echo "Or download from: https://ngrok.com/download"
    echo ""
    exit 1
fi

# Start the server in background
echo "📦 Starting local server..."
cd "$(dirname "$0")"
./start-server.sh &
SERVER_PID=$!

# Wait a moment for server to start
sleep 3

# Start ngrok
echo "🌐 Starting ngrok tunnel..."
echo ""
ngrok http 3000

# When ngrok stops, kill the server
kill $SERVER_PID 2>/dev/null

