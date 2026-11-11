# table2sheet launcher
$ErrorActionPreference = 'SilentlyContinue'
$scriptDir = $PSScriptRoot
if (-not $scriptDir) { $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path }
Set-Location $scriptDir

# Clean up any lingering table2sheet instances on port 8765
$existing = Get-NetTCPConnection -LocalPort 8765 -ErrorAction SilentlyContinue
if ($existing) {
    Stop-Process -Id $existing.OwningProcess -Force -ErrorAction SilentlyContinue
    Start-Sleep -Milliseconds 500
}

# 1. First priority: Check bundled portable runtime (no Python install needed!)
$py = Join-Path $scriptDir 'runtime\python\python.exe'
if (-not (Test-Path $py)) {
    $py = Join-Path $scriptDir '..\runtime\python\python.exe'
}
if (-not (Test-Path $py)) {
    $py = Join-Path $scriptDir 'runtime\python\pythonw.exe'
}
if (-not (Test-Path $py)) {
    $py = Join-Path $scriptDir '..\runtime\python\pythonw.exe'
}

# 2. Second priority: Fall back to system / user Python if runtime not bundled
if (-not (Test-Path $py)) {
    $py = (Get-Command python -ErrorAction SilentlyContinue).Source
}
if (-not $py -or -not (Test-Path $py)) {
    $py = (Get-ChildItem -Path "$env:LOCALAPPDATA\Programs\Python\Python3*\python.exe" -ErrorAction SilentlyContinue | Select-Object -First 1).FullName
}
if (-not $py -or -not (Test-Path $py)) {
    $py = (Get-Command pythonw -ErrorAction SilentlyContinue).Source
}
if (-not $py -or -not (Test-Path $py)) {
    $py = (Get-ChildItem -Path "$env:LOCALAPPDATA\Programs\Python\Python3*\pythonw.exe" -ErrorAction SilentlyContinue | Select-Object -First 1).FullName
}

# Find script
$script = Join-Path $scriptDir 'table2sheet.py'
if (-not (Test-Path $script)) {
    $script = Join-Path $scriptDir 'table2sheet_app\table2sheet.py'
}

if ($py -and (Test-Path $script)) {
    Start-Process -FilePath $py -ArgumentList @("`"$script`"", "--web")
} else {
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show("Could not find table2sheet.py or the Python runtime.", "table2sheet", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
}
