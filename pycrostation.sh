#!/bin/bash

# Check if MsPy-3_11_14 folder exists and has the binary
PYTHON_BINARY="src/linux/MsPy-3_11_14/bin/python3.11"
MSPY_FOLDER="src/linux/MsPy-3_11_14"

if [ ! -d "$MSPY_FOLDER" ] || [ ! -f "$PYTHON_BINARY" ]; then
    echo "MsPy-3_11_14 not found or binary missing. Downloading..."

    # Delete the folder if it exists but is incomplete
    if [ -d "$MSPY_FOLDER" ]; then
        rm -rf "$MSPY_FOLDER"
    fi

    # Create linux directory if it doesn't exist
    mkdir -p src/linux

    # Download the zip file
    wget -O /tmp/MsPy-3_11_14-linux.zip "https://github.com/RisPNG/MsPy/releases/download/3.11.14/MsPy-3_11_14-linux.zip"

    # Extract to src/linux directory
    unzip -q /tmp/MsPy-3_11_14-linux.zip -d src/linux/

    # Clean up zip file
    rm /tmp/MsPy-3_11_14-linux.zip

    # Verify binary exists
    if [ ! -f "$PYTHON_BINARY" ]; then
        echo "Error: Binary not found at $PYTHON_BINARY after extraction"
        exit 1
    fi

    echo "MsPy-3_11_14 successfully downloaded and extracted"
fi

# Check if venv exists and is valid (to determine if we need to run pip install)
VENV_NEWLY_CREATED=false
VENV_NEEDS_RECREATION=false

# Get the absolute path to our Python binary
EXPECTED_PYTHON_HOME=$(cd "$(dirname "$PYTHON_BINARY")" && pwd)

if [ ! -f "src/linux/venv/bin/python" ]; then
    VENV_NEEDS_RECREATION=true
elif [ -f "src/linux/venv/pyvenv.cfg" ]; then
    # Check if venv points to the correct Python home
    if ! grep -q "home = $EXPECTED_PYTHON_HOME" "src/linux/venv/pyvenv.cfg"; then
        echo "Virtual environment points to wrong Python location, recreating..."
        VENV_NEEDS_RECREATION=true
    fi
else
    VENV_NEEDS_RECREATION=true
fi

if [ "$VENV_NEEDS_RECREATION" = true ]; then
    VENV_NEWLY_CREATED=true
    echo "Creating new virtual environment..."

    # Remove existing venv if it exists but is incomplete/wrong platform
    if [ -d "src/linux/venv" ]; then
        rm -rf "src/linux/venv"
    fi

    $PYTHON_BINARY -m venv "src/linux/venv"
fi

# Activate the venv
. src/linux/venv/bin/activate

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
