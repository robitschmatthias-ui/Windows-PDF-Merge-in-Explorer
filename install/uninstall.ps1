# PDF Merge in Explorer - uninstaller
#
# Removes the three "Send to" entries. The project directory (including .venv)
# and the log at %LOCALAPPDATA%\pdf-merge\ are left in place and can be
# deleted by hand afterwards.

$ErrorActionPreference = "Stop"
$SendTo = [Environment]::GetFolderPath("SendTo")

foreach ($name in @("PDF Quick Merge", "PDF Advanced Merge", "Convert to PDF")) {
    $lnk = Join-Path $SendTo "$name.lnk"
    if (Test-Path $lnk) {
        Remove-Item $lnk -Force
        Write-Host "[+] Removed: $lnk"
    } else {
        Write-Host "[ ] Not present: $lnk"
    }
}

Write-Host ""
Write-Host "Not removed (delete by hand if needed):"
Write-Host "  - Project folder: $((Resolve-Path "$PSScriptRoot\..").Path)"
Write-Host "  - Log:            $env:LOCALAPPDATA\pdf-merge\"
