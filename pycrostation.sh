# Create venv only if it doesn't exist
if [ ! -d "src/venv" ]; then
    src/Pynano-3_11_14/bin/python3.11 -m venv "src/venv"
fi

# Activate the venv
. src/venv/bin/activate

# Upgrade pip and install requirements
pip install --upgrade pip
pip install -r requirements.txt

# Run the script
python src/main.py