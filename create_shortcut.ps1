# Creates the "Run table2sheet" shortcut for THIS folder (run once per laptop during setup).
# A Windows shortcut stores an absolute path, so it must be created on each laptop
# after the app folder is in its final location.
#
# Usage (from the app folder):
#   powershell -NoProfile -ExecutionPolicy Bypass -File create_shortcut.ps1
#   powershell -NoProfile -ExecutionPolicy Bypass -File create_shortcut.ps1 -Desktop
param([switch]$Desktop)

$appDir = $PSScriptRoot
if (-not $appDir) { $appDir = Split-Path -Parent $MyInvocation.MyCommand.Path }
$launcher = Join-Path $appDir 'launch.ps1'
if (-not (Test-Path $launcher)) { Write-Error "launch.ps1 not found in $appDir"; exit 1 }

$psExe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
$shell = New-Object -ComObject WScript.Shell

$targets = @(Join-Path $appDir 'Run table2sheet.lnk')
if ($Desktop) { $targets += Join-Path ([Environment]::GetFolderPath('Desktop')) 'Run table2sheet.lnk' }

foreach ($path in $targets) {
    $lnk = $shell.CreateShortcut($path)
    $lnk.TargetPath = $psExe
    $lnk.Arguments = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$launcher`""
    $lnk.WorkingDirectory = $appDir
    $lnk.IconLocation = "$psExe,0"
    $lnk.Description = 'Start table2sheet (opens http://localhost:8765)'
    $lnk.Save()
    Write-Host "Created: $path"
}
