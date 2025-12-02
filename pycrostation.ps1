param([switch]$Update)

# Use the actual location of the script, even if run from shortcuts or another user account
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition
Set-Location $scriptPath

$ErrorActionPreference = 'Stop'

# Check if MsPy-3_11_14 folder exists and has the binary
$pythonBinary = "src\win\MsPy-3_11_14\python.exe"
$mspyFolder = "src\win\MsPy-3_11_14"

if (-not (Test-Path $mspyFolder) -or -not (Test-Path $pythonBinary)) {
    Write-Host "MsPy-3_11_14 not found or binary missing. Downloading..."

    # Delete the folder if it exists but is incomplete
    if (Test-Path $mspyFolder) {
        Remove-Item $mspyFolder -Recurse -Force
    }

    # Create win directory if it doesn't exist
    New-Item -ItemType Directory -Force -Path "src\win" | Out-Null

    # Download the zip file
    $zipPath = Join-Path $env:TEMP "MsPy-3_11_14-win.zip"
    Invoke-WebRequest -Uri "https://github.com/RisPNG/MsPy/releases/download/3.11.14/MsPy-3_11_14-win.zip" -OutFile $zipPath

    # Extract to src/win directory
    Expand-Archive -Path $zipPath -DestinationPath "src\win\" -Force

    # Clean up zip file
    Remove-Item $zipPath

    # Verify binary exists
    if (-not (Test-Path $pythonBinary)) {
        Write-Error "Error: Binary not found at $pythonBinary after extraction"
        exit 1
    }

    Write-Host "MsPy-3_11_14 successfully downloaded and extracted"
}

# Check if venv exists and is valid (to determine if we need to run pip install)
$venvNewlyCreated = $false
$venvNeedsRecreation = $false

# Get the absolute path to our Python binary directory
$expectedPythonHome = Split-Path -Parent (Resolve-Path $pythonBinary)

if (-not (Test-Path "src\win\venv\Scripts\python.exe")) {
    $venvNeedsRecreation = $true
} elseif (Test-Path "src\win\venv\pyvenv.cfg") {
    # Check if venv points to the correct Python home
    $pyvenvContent = Get-Content "src\win\venv\pyvenv.cfg" -Raw
    if ($pyvenvContent -notmatch [regex]::Escape("home = $expectedPythonHome")) {
        Write-Host "Virtual environment points to wrong Python location, recreating..."
        $venvNeedsRecreation = $true
    }
} else {
    $venvNeedsRecreation = $true
}

if ($venvNeedsRecreation) {
    $venvNewlyCreated = $true
    Write-Host "Creating new virtual environment..."

    # Remove existing venv if it exists but is incomplete/wrong platform
    if (Test-Path "src\win\venv") {
        Remove-Item "src\win\venv" -Recurse -Force
    }

    & $pythonBinary -m venv "src\win\venv"
}

$venvPy = Join-Path $scriptPath "src\win\venv\Scripts\python.exe"

# Only run pip install if venv was newly created
if ($venvNewlyCreated) {
    Write-Host "Installing dependencies..."
    & $venvPy -m pip install --upgrade pip
    & $venvPy -m pip install -r requirements.txt
} else {
    Write-Host "Using existing virtual environment (skipping pip install)"
}

if ($Update -and (Test-Path ".pycro-repo")) {
    Remove-Item ".pycro-repo" -Recurse -Force
}

$env:PYCRO_REPO_URL    = "https://github.com/RisPNG/pycro-station.git"
$env:PYCRO_REPO_BRANCH = "main"
$env:PYCRO_REPO_SUBDIR = "pycros"

& $venvPy "src\main.py"
