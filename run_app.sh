#!/bin/bash

# Script to run the A&H Brass Shopify Feed Generator Streamlit app

echo "A&H Brass Shopify Feed Generator"
echo "--------------------------------"

# Check if Python is installed
if ! command -v python3 &> /dev/null; then
    echo "Error: Python 3 is required but not installed."
    exit 1
fi

# Check if pip is installed
if ! command -v pip3 &> /dev/null; then
    echo "Error: pip3 is required but not installed."
    exit 1
fi

# Check if requirements.txt exists
if [ ! -f "requirements.txt" ]; then
    echo "Error: requirements.txt not found."
    exit 1
fi

# Install dependencies
echo "Installing dependencies..."
pip3 install -r requirements.txt

# Check if app.py exists
if [ ! -f "app.py" ]; then
    echo "Error: app.py not found."
    exit 1
fi

# Run the app
echo "Starting Streamlit app..."
streamlit run app.py

echo "App closed." 