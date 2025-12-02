#!/bin/bash

# Check if MsPy-3_11_14 folder exists and has the binary
PYTHON_BINARY="src/MsPy-3_11_14/bin/python3.11"
MSPY_FOLDER="src/MsPy-3_11_14"

if [ ! -d "$MSPY_FOLDER" ] || [ ! -f "$PYTHON_BINARY" ]; then
    echo "MsPy-3_11_14 not found or binary missing. Downloading..."

    # Delete the folder if it exists but is incomplete
    if [ -d "$MSPY_FOLDER" ]; then
        rm -rf "$MSPY_FOLDER"
    fi

    # Download the zip file
    wget -O /tmp/MsPy-3_11_14-linux.zip "https://github.com/RisPNG/MsPy/releases/download/3.11.14/MsPy-3_11_14-linux.zip"

    # Extract to src directory
    unzip -q /tmp/MsPy-3_11_14-linux.zip -d src/

    # Clean up zip file
    rm /tmp/MsPy-3_11_14-linux.zip

    # Verify binary exists
    if [ ! -f "$PYTHON_BINARY" ]; then
        echo "Error: Binary not found at $PYTHON_BINARY after extraction"
        exit 1
    fi

    echo "MsPy-3_11_14 successfully downloaded and extracted"
fi

# Check if venv exists (to determine if we need to run pip install)
VENV_NEWLY_CREATED=false

if [ ! -d "src/venv" ]; then
    VENV_NEWLY_CREATED=true
    echo "Creating new virtual environment..."
    $PYTHON_BINARY -m venv "src/venv"
fi

# Activate the venv
. src/venv/bin/activate

# Only run pip install if venv was newly created
if [ "$VENV_NEWLY_CREATED" = true ]; then
    echo "Installing dependencies..."
    pip install --upgrade pip
    pip install -r requirements.txt
else
    echo "Using existing virtual environment (skipping pip install)"
fi

# Run the script
python src/main.py
