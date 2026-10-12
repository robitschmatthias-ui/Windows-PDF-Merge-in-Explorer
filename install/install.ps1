# PDF Merge in Explorer - Windows installer
#
#   - Checks for Python 3.10+ and installs it via winget if missing
#   - Detects and repairs a broken .venv from an earlier run
#   - Creates a Python venv and installs the dependencies
#   - Creates three entries in the Explorer "Send to" menu:
#       * PDF Quick Merge
#       * PDF Advanced Merge
#       * Convert to PDF
#
# Start by double-clicking install\install.bat (bypasses the ExecutionPolicy).
#
# Note: keep this file ASCII-only - Windows PowerShell 5.1 reads scripts
# without BOM as ANSI and would mangle non-ASCII characters.

$ErrorActionPreference = "Stop"

$ProjectDir       = (Resolve-Path "$PSScriptRoot\..").Path
$VenvDir          = Join-Path $ProjectDir ".venv"
$PythonExe        = Join-Path $VenvDir "Scripts\python.exe"
$PythonwExe       = Join-Path $VenvDir "Scripts\pythonw.exe"
$RequirementsFile = Join-Path $ProjectDir "requirements.txt"
$MainScript       = Join-Path $ProjectDir "src\pdf_merge.py"
$IconFile         = Join-Path $ProjectDir "pdf-merge-icon.ico"

Write-Host "PDF Merge in Explorer - Installer (Windows)"
Write-Host "==========================================="
Write-Host "Project directory: $ProjectDir"
Write-Host ""

# --- Helper: run a step with a spinner ---
function Invoke-Step {
    param(
        [string]   $Label,
        [string]   $Exe,
        [string[]] $Arguments = @()
    )
    $frames  = @('/', '-', '\', '|')
    $outFile = [System.IO.Path]::GetTempFileName()
    $errFile = [System.IO.Path]::GetTempFileName()

    $quotedArgs = $Arguments | ForEach-Object {
        if ($_ -match '\s') { "`"$_`"" } else { $_ }
    }
    $proc = Start-Process -FilePath $Exe -ArgumentList $quotedArgs `
        -RedirectStandardOutput $outFile -RedirectStandardError $errFile `
        -NoNewWindow -PassThru

    $i = 0
    while (-not $proc.HasExited) {
        Write-Host -NoNewline "`r[$($frames[$i % 4])] $Label ..."
        $i++
        Start-Sleep -Milliseconds 120
    }
    $proc.WaitForExit()

    if (([int]$proc.ExitCode) -ne 0) {
        Write-Host "`r[X] $Label (Exit $($proc.ExitCode))                         "
        $stdout = Get-Content $outFile -Raw -ErrorAction SilentlyContinue
        $stderr = Get-Content $errFile -Raw -ErrorAction SilentlyContinue
        if ($stderr) {
            Write-Host "    --- stderr ---"
            $stderr.TrimEnd() -split "`r?`n" | ForEach-Object { Write-Host "    $_" }
        }
        if ($stdout) {
            Write-Host "    --- stdout ---"
            $stdout.TrimEnd() -split "`r?`n" | ForEach-Object { Write-Host "    $_" }
        }
        if (-not $stderr -and -not $stdout) {
            Write-Host "    (no output)"
            Write-Host "    Command: $Exe $($Arguments -join ' ')"
        }
        Remove-Item $outFile, $errFile -Force -ErrorAction SilentlyContinue
        exit 1
    }
    Write-Host "`r[+] $Label                              "
    Remove-Item $outFile, $errFile -Force -ErrorAction SilentlyContinue
}

# --- Python 3.10+ ---
# Prefers the 'py' launcher; skips the Microsoft Store stub at
# WindowsApps\python.exe (it hangs the shell).
$Script:PythonCmd = $null

function Test-PythonVersion {
    param([string]$Exe, [string[]]$PreArgs = @())
    try {
        $allArgs = $PreArgs + @("-c", "import sys; print(f'{sys.version_info.major}.{sys.version_info.minor}')")
        $ver = & $Exe @allArgs 2>$null
        if (-not $ver) { return $false }
        $parts = $ver.ToString().Trim().Split('.')
        $major = [int]$parts[0]
        $minor = [int]$parts[1]
        return (($major -gt 3) -or ($major -eq 3 -and $minor -ge 10))
    } catch {
        return $false
    }
}

function Test-PythonOk {
    $candidates = @(
        @{ Exe = "py";      PreArgs = @("-3") },
        @{ Exe = "python";  PreArgs = @() },
        @{ Exe = "python3"; PreArgs = @() }
    )
    foreach ($c in $candidates) {
        $cmd = Get-Command $c.Exe -ErrorAction SilentlyContinue
        if (-not $cmd) { continue }
        if ($cmd.Source -and $cmd.Source -like "*\WindowsApps\*") { continue }
        if (Test-PythonVersion -Exe $c.Exe -PreArgs $c.PreArgs) {
            $Script:PythonCmd = @($c.Exe) + $c.PreArgs
            return $true
        }
    }
    # Fallback: standard install locations (in case winget installed Python
    # but this shell still has the old PATH)
    $roots = @("$env:LocalAppData\Programs\Python", "$env:ProgramFiles", "${env:ProgramFiles(x86)}")
    foreach ($root in $roots) {
        if (-not (Test-Path $root)) { continue }
        $found = Get-ChildItem -Path $root -Directory -Filter "Python3*" -ErrorAction SilentlyContinue |
            Sort-Object Name -Descending
        foreach ($dir in $found) {
            $p = Join-Path $dir.FullName "python.exe"
            if ((Test-Path $p) -and (Test-PythonVersion -Exe $p)) {
                $Script:PythonCmd = @($p)
                return $true
            }
        }
    }
    return $false
}

if (Test-PythonOk) {
    Write-Host "[+] Python found ($($Script:PythonCmd -join ' '))"
} else {
    if ($env:PDF_MERGE_INSTALLER_RELAUNCHED -eq "1") {
        Write-Host "[X] Python was installed but still cannot be found."
        Write-Host "    Please check in a new PowerShell:  py --version"
        Write-Host "    Otherwise install Python manually: https://www.python.org/downloads/"
        Write-Host "    IMPORTANT: enable 'Add Python to PATH'."
        exit 1
    }

    Write-Host "[>] Python 3.10+ not found. Installing via winget..."
    $winget = Get-Command winget -ErrorAction SilentlyContinue
    if (-not $winget) {
        Write-Host "[X] winget not available."
        Write-Host "    Please install Python 3.10+ manually: https://www.python.org/downloads/"
        Write-Host "    IMPORTANT: enable 'Add Python to PATH'."
        exit 1
    }
    Write-Host "    (This can take a minute - winget is downloading Python...)"
    winget install --id Python.Python.3.12 -e --silent --accept-package-agreements --accept-source-agreements

    $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" +
                [System.Environment]::GetEnvironmentVariable("Path", "User")
    $env:PDF_MERGE_INSTALLER_RELAUNCHED = "1"

    Write-Host "[+] Python installed. Continuing the installer in a new session..."
    Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Wait
    exit 0
}

# --- venv ---
# A venv counts as broken if python.exe is missing or does not start
# (e.g. after an aborted run or a removed base Python).
if (Test-Path $VenvDir) {
    $venvOk = $false
    if (Test-Path $PythonExe) {
        & $PythonExe -c "import sys" 2>$null | Out-Null
        if (([int]$LASTEXITCODE) -eq 0) { $venvOk = $true }
    }
    if (-not $venvOk) {
        Write-Host "[>] Removing broken venv from an earlier run..."
        Remove-Item -Recurse -Force $VenvDir
    }
}

if (-not (Test-Path $VenvDir)) {
    $pyExec = $Script:PythonCmd[0]
    $pyArgs = if ($Script:PythonCmd.Length -gt 1) {
        $Script:PythonCmd[1..($Script:PythonCmd.Length - 1)] + @("-m", "venv", $VenvDir)
    } else {
        @("-m", "venv", $VenvDir)
    }
    Invoke-Step "Creating Python venv" $pyExec $pyArgs
} else {
    Write-Host "[+] Python venv present"
}

Invoke-Step "Updating pip" `
    $PythonExe @("-m", "pip", "install", "--quiet", "--upgrade", "pip")

Invoke-Step "Installing Python dependencies (this can take a moment)" `
    $PythonExe @("-m", "pip", "install", "--quiet", "--upgrade", "-r", $RequirementsFile)

# Check that all modules actually load (including the pywin32 DLLs)
Invoke-Step "Verifying installation" `
    $PythonExe @("-c", "import pypdf, img2pdf, PIL, win32com.client, pythoncom, tkinter")

# --- Office / LibreOffice (information only, nothing is installed) ---
$hasWord = Test-Path "Registry::HKEY_CLASSES_ROOT\Word.Application\CLSID"
$soffice = @("$env:ProgramFiles\LibreOffice\program\soffice.exe",
             "${env:ProgramFiles(x86)}\LibreOffice\program\soffice.exe") |
    Where-Object { Test-Path $_ } | Select-Object -First 1
if ($hasWord) {
    Write-Host "[+] Microsoft Office found - Office files will be converted with it"
} elseif ($soffice) {
    Write-Host "[+] LibreOffice found - Office files will be converted with it"
} else {
    Write-Host "[!] Neither Microsoft Office nor LibreOffice found."
    Write-Host "    PDFs and images work; Word/Excel/PowerPoint files need"
    Write-Host "    one of the two."
}

# --- "Send to" entries ---
$SendTo = [Environment]::GetFolderPath("SendTo")
$Entries = @(
    @{ Name = "PDF Quick Merge";    Mode = "--quick";    Desc = "Merge files straight into Merge.pdf" },
    @{ Name = "PDF Advanced Merge"; Mode = "--advanced"; Desc = "Choose order, page size and orientation" },
    @{ Name = "Convert to PDF";     Mode = "--convert";  Desc = "Convert each file to its own PDF in the same folder" }
)

try {
    $shell = New-Object -ComObject WScript.Shell
    foreach ($e in $Entries) {
        $lnkPath = Join-Path $SendTo "$($e.Name).lnk"
        $lnk = $shell.CreateShortcut($lnkPath)
        $lnk.TargetPath       = $PythonwExe
        $lnk.Arguments        = "`"$MainScript`" $($e.Mode)"
        $lnk.WorkingDirectory = $ProjectDir
        $lnk.IconLocation     = "$IconFile,0"
        $lnk.Description      = $e.Desc
        $lnk.Save()
        Write-Host "[+] Send to: $($e.Name)"
    }
} catch {
    Write-Host "[X] Could not create the Send to entries"
    Write-Host "    Error:  $($_.Exception.Message)"
    Write-Host "    Folder: $SendTo"
    exit 1
}

Write-Host ""
Write-Host "[+] Installation complete."
Write-Host "    Select files > right-click > Send to > PDF Quick Merge / PDF Advanced Merge / Convert to PDF"
Write-Host "    (Windows 11: first 'Show more options', or Shift + right-click)"
