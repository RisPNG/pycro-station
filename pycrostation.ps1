# Create venv only if it doesn't exist
if (-not (Test-Path "src\venv-win")) {
    src\MsPy-3_11_14-win\python.exe -m venv "src\venv-win"
}

# Activate the venv
& src\venv-win\Scripts\Activate.ps1

# Upgrade pip and install requirements
python.exe -m pip install --upgrade pip
pip install -r requirements.txt

# Run the script
python src\main.py