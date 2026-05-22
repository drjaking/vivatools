# start_viva_console.ps1
# ==============================================================================
# Helper script to run the DClinPsy Viva Examiner Matching Console locally.
# This script starts PostgreSQL and the FastAPI/HTMX dev server in independent
# background processes so they persist even when the AI agent container restarts.
# ==============================================================================

Write-Host "==================================================" -ForegroundColor Cyan
Write-Host " Starting DClinPsy Viva Matching Console Services" -ForegroundColor Cyan
Write-Host "==================================================" -ForegroundColor Cyan

# 1. Start PostgreSQL Database Server if not running
$pgPort = 5432
$pgSocket = Get-NetTCPConnection -LocalPort $pgPort -ErrorAction SilentlyContinue

if ($pgSocket) {
    Write-Host "[*] PostgreSQL is already running on port $pgPort." -ForegroundColor Green
} else {
    Write-Host "[+] Starting PostgreSQL server using pg_ctl..." -ForegroundColor Yellow
    
    # Run pg_ctl start with -w (wait) so it blocks until the database is fully ready
    # This prints any startup errors directly to this console window instead of hiding them.
    & "C:\Program Files\PostgreSQL\17\bin\pg_ctl.exe" start -D "C:\Program Files\PostgreSQL\17\data" -w
    
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Failed to start PostgreSQL server. Please check the error above."
        Exit 1
    }
    Write-Host "[*] PostgreSQL is ready." -ForegroundColor Green
}

# 2. Start Uvicorn FastAPI Server if not running
$webPort = 8000
$webSocket = Get-NetTCPConnection -LocalPort $webPort -ErrorAction SilentlyContinue

if ($webSocket) {
    Write-Host "[*] FastAPI/Uvicorn is already running on port $webPort." -ForegroundColor Green
} else {
    Write-Host "[+] Starting FastAPI/Uvicorn web console in a new window..." -ForegroundColor Yellow
    
    # Start Uvicorn server in a separate PowerShell window so logs are visible
    Start-Process -FilePath "powershell.exe" `
                  -ArgumentList "-NoExit", "-Command", "Write-Host 'Starting Uvicorn Server...' -ForegroundColor Cyan; .venv\Scripts\python.exe -m uvicorn app.main:app --host 127.0.0.1 --port $webPort --reload" `
                  -WorkingDirectory "C:\vivatools" `
                  -WindowStyle Normal
                  
    Start-Sleep -Seconds 2
}

# 3. Open the web browser to the matching console
Write-Host "[+] Opening DClinPsy Viva Matching Console in your browser..." -ForegroundColor Green
Start-Process "http://127.0.0.1:8000/import"

Write-Host "==================================================" -ForegroundColor Cyan
Write-Host " Ready! Close the Uvicorn terminal window to stop." -ForegroundColor Cyan
Write-Host "==================================================" -ForegroundColor Cyan
