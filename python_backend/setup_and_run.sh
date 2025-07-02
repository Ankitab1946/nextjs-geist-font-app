#!/bin/bash

echo "=== Fuzzy Column Matcher Setup and Run Script ==="

# Check if Python is installed
if ! command -v python3 &> /dev/null; then
    echo "Python 3 is not installed. Please install Python 3 and try again."
    exit 1
fi

# Check if pip is installed
if ! command -v pip3 &> /dev/null; then
    echo "pip3 is not installed. Please install pip3 and try again."
    exit 1
fi

# Create virtual environment
echo "Creating virtual environment..."
python3 -m venv venv

# Activate virtual environment
echo "Activating virtual environment..."
source venv/bin/activate

# Install requirements
echo "Installing requirements..."
pip install -r requirements.txt

# Run tests
echo "Running tests..."
python test_matcher.py

# Start Streamlit app
echo "Starting Streamlit app..."
streamlit run streamlit_app.py

# Note: The script will keep running until the Streamlit app is closed
# The virtual environment will remain active until you run 'deactivate'
