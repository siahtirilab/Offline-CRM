$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $projectRoot

$appName = "OfflineCRM"
$version = "2"
$desktopIcon = Join-Path $env:USERPROFILE "Desktop\Icon.png"
$projectIconPng = Join-Path $projectRoot "Icon.png"
$projectIconIco = Join-Path $projectRoot "Icon.ico"

if (-not (Test-Path $projectIconPng)) {
    if (Test-Path $desktopIcon) {
        Copy-Item $desktopIcon $projectIconPng -Force
    }
    else {
        throw "Icon.png was not found in the project root or Desktop."
    }
}

python -c "from PIL import Image; img=Image.open('Icon.png').convert('RGBA'); img.save('Icon.ico', format='ICO', sizes=[(16,16),(24,24),(32,32),(48,48),(64,64),(128,128),(256,256)])"
python -m pip install -r requirements.txt
python -m pip install pyinstaller

if (Test-Path 'build') { Remove-Item 'build' -Recurse -Force }
if (Test-Path \"dist\$appName\") { Remove-Item \"dist\$appName\" -Recurse -Force }
if (Test-Path 'output') { Remove-Item 'output' -Recurse -Force }
New-Item -ItemType Directory -Path 'output' -Force | Out-Null

python -m PyInstaller `
  --noconfirm `
  --clean `
  --windowed `
  --name $appName `
  --icon $projectIconIco `
  --add-data "Icon.png;." `
  --hidden-import pystray._win32 `
  --collect-submodules winotify `
  --collect-submodules ttkbootstrap `
  main.py

$nsisCandidates = @(
    "${env:ProgramFiles(x86)}\NSIS\makensis.exe",
    "${env:ProgramFiles}\NSIS\makensis.exe"
)

$makensis = $nsisCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $makensis) {
    throw "makensis.exe was not found. Install NSIS first."
}

& $makensis "installer.nsi"

Write-Host ""
Write-Host "Build complete."
Write-Host "EXE folder: $projectRoot\dist\$appName"
Write-Host "Installer folder: $projectRoot\output"
