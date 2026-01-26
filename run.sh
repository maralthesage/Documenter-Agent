#!/bin/bash

# Excel Documentation AI Agent - Run Script
# Starts the Streamlit application

echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo "Excel Documentation AI Agent"
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo ""

# Check if .env file exists
if [ ! -f .env ]; then
    echo "⚠️  .env file not found. Creating from .env.example..."
    cp .env.example .env
    echo "✓ Created .env file"
fi

echo ""
echo "Checking prerequisites..."
echo ""

# Check if Ollama is running
echo -n "Checking Ollama connection... "
if curl -s http://localhost:11434/api/tags > /dev/null 2>&1; then
    echo "✓ Connected"
else
    echo "✗ Not connected"
    echo ""
    echo "Please start Ollama with: ollama serve"
    exit 1
fi

# Check if model is available
MODEL=${OLLAMA_MODEL:-mistral}
echo -n "Checking for $MODEL model... "
MODELS=$(curl -s http://localhost:11434/api/tags | grep -o '"name":"[^"]*' | cut -d'"' -f4)
if echo "$MODELS" | grep -q "$MODEL"; then
    echo "✓ Found"
else
    echo "✗ Not found"
    echo ""
    echo "Please pull the model: ollama pull $MODEL"
    exit 1
fi

echo ""
echo "Starting Streamlit application..."
echo ""
echo "🌐 Open your browser to: http://localhost:8501"
echo ""

streamlit run app.py
