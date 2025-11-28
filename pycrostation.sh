# Create venv only if it doesn't exist
if [ ! -d "src/venv" ]; then
    src/MsPy-3_11_14-linux/bin/python3.11 -m venv "src/venv-linux"
fi

# Activate the venv
. src/venv-linux/bin/activate

# Upgrade pip and install requirements
pip install --upgrade pip
pip install -r requirements.txt

# Run the script
python src/main.py