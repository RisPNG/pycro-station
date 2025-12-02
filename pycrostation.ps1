param([switch]$Update)

# Use the actual location of the script, even if run from shortcuts or another user account
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition
Set-Location $scriptPath

$ErrorActionPreference = 'Stop'

if (-not (Test-Path "src\venv-win\Scripts\python.exe")) {
    & "src\MsPy-3_11_14-win\python.exe" -m venv "src\venv-win"
    $venvPy = Join-Path $scriptPath "src\venv-win\Scripts\python.exe"

    & $venvPy -m pip install --upgrade pip
    & $venvPy -m pip install -r requirements.txt
}

$venvPy = Join-Path $scriptPath "src\venv-win\Scripts\python.exe"

if ($Update -and (Test-Path ".pycro-repo")) {
    Remove-Item ".pycro-repo" -Recurse -Force
}

$env:PYCRO_REPO_URL    = "https://github.com/RisPNG/pycro-station.git"
$env:PYCRO_REPO_BRANCH = "main"
$env:PYCRO_REPO_SUBDIR = "pycros"

& $venvPy "src\main.py"
