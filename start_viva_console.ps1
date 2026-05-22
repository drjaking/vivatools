# start_viva_console.ps1
# ==============================================================================
# Helper script to run the DClinPsy Viva Examiner Matching Console locally.
# This script starts PostgreSQL and the FastAPI/HTMX dev server in independent
# background processes so they persist even when the AI agent container restarts.
# ==============================================================================

Write-Host "==================================================" -ForegroundColor Cyan
Write-Host " Starting DClinPsy Viva Matching Console Services" -ForegroundColor Cyan
Write-Host "==================================================" -ForegroundColor Cyan

# Determine python virtual environment folder (venv or .venv)
$venvDir = "venv"
if (Test-Path "$PSScriptRoot\.venv") {
    $venvDir = ".venv"
}

# 1. Start PostgreSQL Database Server if not running
$pgPort = 5432
$pgSocket = Get-NetTCPConnection -LocalPort $pgPort -ErrorAction SilentlyContinue | Where-Object { $_.LocalAddress -in '127.0.0.1', '::1', '0.0.0.0' }

if ($pgSocket) {
    Write-Host "[*] PostgreSQL is already running on port $pgPort." -ForegroundColor Green
} else {
    Write-Host "[+] Detecting PostgreSQL installation..." -ForegroundColor Yellow
    
    $pgBaseDir = "C:\Program Files\PostgreSQL"
    if (-not (Test-Path $pgBaseDir)) {
        Write-Error "PostgreSQL directory '$pgBaseDir' not found. Please install PostgreSQL."
        Exit 1
    }
    
    # Get the latest version folder (e.g. 18, 17, etc.)
    $pgVerFolder = Get-ChildItem $pgBaseDir -Directory | Sort-Object Name -Descending | Select-Object -First 1
    if (-not $pgVerFolder) {
        Write-Error "No PostgreSQL version folder found in '$pgBaseDir'."
        Exit 1
    }
    
    $pgVersion = $pgVerFolder.Name
    $pgCtl = Join-Path $pgVerFolder.FullName "bin\pg_ctl.exe"
    $pgData = Join-Path $pgVerFolder.FullName "data"
    
    Write-Host "[*] Found PostgreSQL version $pgVersion at: $($pgVerFolder.FullName)" -ForegroundColor Green
    Write-Host "[+] Starting PostgreSQL server using pg_ctl..." -ForegroundColor Yellow
    
    # Run pg_ctl start with -w (wait) so it blocks until the database is fully ready
    # This prints any startup errors directly to this console window instead of hiding them.
    & $pgCtl start -D $pgData -w
    
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Failed to start PostgreSQL server. Please check the error above."
        Exit 1
    }
    Write-Host "[*] PostgreSQL is ready." -ForegroundColor Green
}

# 2. Start Uvicorn FastAPI Server if not running
$webPort = 8000
$webSocket = Get-NetTCPConnection -LocalPort $webPort -ErrorAction SilentlyContinue | Where-Object { $_.LocalAddress -in '127.0.0.1', '::1', '0.0.0.0' }

if ($webSocket) {
    Write-Host "[*] FastAPI/Uvicorn is already running on port $webPort." -ForegroundColor Green
} else {
    Write-Host "[+] Starting FastAPI/Uvicorn web console in a new window..." -ForegroundColor Yellow
    
    # Start Uvicorn server in a separate PowerShell window so logs are visible
    Start-Process -FilePath "powershell.exe" `
                  -ArgumentList "-NoExit", "-Command", "Write-Host 'Starting Uvicorn Server...' -ForegroundColor Cyan; $venvDir\Scripts\python.exe -m uvicorn app.main:app --host 127.0.0.1 --port $webPort --reload" `
                  -WorkingDirectory $PSScriptRoot `
                  -WindowStyle Normal
                  
    Start-Sleep -Seconds 2
}

# 3. Open the web browser to the matching console
Write-Host "[+] Opening DClinPsy Viva Matching Console in your browser..." -ForegroundColor Green
Start-Process "http://127.0.0.1:8000/import"

Write-Host "==================================================" -ForegroundColor Cyan
Write-Host " Ready! Close the Uvicorn terminal window to stop." -ForegroundColor Cyan
Write-Host "==================================================" -ForegroundColor Cyan

