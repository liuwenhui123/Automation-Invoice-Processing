$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$scriptPath = Join-Path $projectRoot "invoice_renamer.py"

python -m pip install --upgrade pip
python -m pip install --upgrade pyinstaller pypdf openpyxl

python -OO -m PyInstaller `
  --noconfirm `
  --clean `
  --onefile `
  --console `
  --name "invoice_renamer" `
  --exclude-module test `
  --exclude-module unittest `
  --exclude-module pdb `
  --exclude-module PIL `
  --exclude-module numpy `
  --exclude-module cryptography `
  --exclude-module Crypto `
  --exclude-module psutil `
  --exclude-module lxml `
  --exclude-module matplotlib `
  --exclude-module IPython `
  --exclude-module jedi `
  --exclude-module pygments `
  --exclude-module pkg_resources `
  --exclude-module setuptools `
  --exclude-module html5lib `
  --exclude-module scipy `
  --exclude-module pytest `
  --exclude-module pyparsing `
  $scriptPath

Write-Host ""
Write-Host "Build complete: $(Join-Path $projectRoot 'dist\\invoice_renamer.exe')"
