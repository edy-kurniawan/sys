# ========================================
# INSTALL MONTHLY TASK SCHEDULER
# ========================================
# Compatible: Windows 7 (PS 2.0) - Windows 11
# Schedule: First week of month (1-7), once per month, with retry
# Requirement: Internet connection

param(
    [string]$ExePath = "$PSScriptRoot\bin\Release\MaintenanceApp.exe",
    [string]$TaskName = "PC Maintenance Monthly Report",
    [string]$Time = "12:00"
)

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Installing Monthly Task Scheduler" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

# Check if running as Administrator
$isAdmin = $false
try {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    $isAdmin = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
} catch {
    # Fallback for PS 2.0
    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

if (-not $isAdmin) {
    Write-Host "[ERROR] This script must be run as Administrator!" -ForegroundColor Red
    Write-Host "Right-click PowerShell and select 'Run as Administrator'" -ForegroundColor Yellow
    pause
    exit 1
}

# Check if EXE exists
if (-not (Test-Path $ExePath)) {
    Write-Host "[ERROR] MaintenanceApp.exe not found at: $ExePath" -ForegroundColor Red
    Write-Host "Please build the project first using build.bat" -ForegroundColor Yellow
    pause
    exit 1
}

Write-Host "[INFO] EXE Path: $ExePath" -ForegroundColor Green
Write-Host "[INFO] Task Name: $TaskName" -ForegroundColor Green
Write-Host "[INFO] Schedule: First week of month (1-7) at $Time" -ForegroundColor Green
Write-Host "[INFO] Mode: Run once per month with retry`n" -ForegroundColor Green

# Create wrapper script directory
$wrapperDir = Join-Path $PSScriptRoot "TaskScheduler"
if (-not (Test-Path $wrapperDir)) {
    New-Item -Path $wrapperDir -ItemType Directory -Force | Out-Null
}

$wrapperScript = Join-Path $wrapperDir "MonthlyMaintenanceWrapper.ps1"
$markerDir = Join-Path $wrapperDir "Markers"

# Create marker directory
if (-not (Test-Path $markerDir)) {
    New-Item -Path $markerDir -ItemType Directory -Force | Out-Null
}

Write-Host "[CREATE] Generating wrapper script..." -ForegroundColor Yellow

# Create wrapper script with PS 2.0 compatibility
$wrapperContent = @'
# ========================================
# MONTHLY MAINTENANCE WRAPPER SCRIPT
# ========================================
# Compatible: PowerShell 2.0+
# Purpose: Run maintenance once per month with internet check and retry

param(
    [string]$ExePath = "{0}",
    [string]$MarkerDir = "{1}"
)
'@ -f $ExePath, $markerDir

$wrapperContent += @'


# Function to check internet connectivity (PS 2.0 compatible)
function Test-InternetConnection {
    try {
        # Method 1: Try ping Google DNS
        $ping = New-Object System.Net.NetworkInformation.Ping
        $result = $ping.Send("8.8.8.8", 1000)
        if ($result.Status -eq "Success") {
            return $true
        }
    } catch {}
    
    try {
        # Method 2: Try HTTP request
        $request = [System.Net.WebRequest]::Create("http://www.msftconnecttest.com/connecttest.txt")
        $request.Timeout = 3000
        $response = $request.GetResponse()
        $response.Close()
        return $true
    } catch {}
    
    return $false
}

# Function to write log (PS 2.0 compatible)
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # Console output
    switch ($Level) {
        "ERROR" { Write-Host $logMessage -ForegroundColor Red }
        "SUCCESS" { Write-Host $logMessage -ForegroundColor Green }
        "WARNING" { Write-Host $logMessage -ForegroundColor Yellow }
        default { Write-Host $logMessage -ForegroundColor White }
    }
    
    # File log
    $logFile = Join-Path $MarkerDir "maintenance.log"
    Add-Content -Path $logFile -Value $logMessage -ErrorAction SilentlyContinue
}

# Main execution
Write-Log "========================================" "INFO"
Write-Log "Monthly Maintenance Task Started" "INFO"
Write-Log "========================================" "INFO"

# Get current month marker
$currentMonth = Get-Date -Format "yyyy-MM"
$markerFile = Join-Path $MarkerDir "$currentMonth.completed"

# Check if already run this month
if (Test-Path $markerFile) {
    Write-Log "Task already completed this month ($currentMonth)" "WARNING"
    Write-Log "Marker file: $markerFile" "INFO"
    Write-Log "Skipping execution." "WARNING"
    exit 0
}

# Check current date (must be 1-7)
$currentDay = (Get-Date).Day
if ($currentDay -lt 1 -or $currentDay -gt 7) {
    Write-Log "Current date is $currentDay - outside first week window (1-7)" "WARNING"
    Write-Log "Task will retry tomorrow if in valid date range" "INFO"
    exit 0
}

Write-Log "Current date: $currentDay (within first week window)" "INFO"

# Check internet connection
Write-Log "Checking internet connectivity..." "INFO"
if (-not (Test-InternetConnection)) {
    Write-Log "No internet connection detected" "ERROR"
    Write-Log "Task will retry when internet is available" "WARNING"
    exit 1
}

Write-Log "Internet connection: OK" "SUCCESS"

# Check if EXE exists
if (-not (Test-Path $ExePath)) {
    Write-Log "Maintenance executable not found: $ExePath" "ERROR"
    exit 1
}

# Run maintenance application
Write-Log "Starting maintenance application..." "INFO"
Write-Log "EXE: $ExePath" "INFO"

try {
    $startTime = Get-Date
    
    # Start process and wait for completion
    $processInfo = New-Object System.Diagnostics.ProcessStartInfo
    $processInfo.FileName = $ExePath
    $processInfo.WorkingDirectory = Split-Path $ExePath
    $processInfo.UseShellExecute = $false
    $processInfo.RedirectStandardOutput = $true
    $processInfo.RedirectStandardError = $true
    $processInfo.CreateNoWindow = $true
    
    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $processInfo
    
    $stdout = New-Object System.Text.StringBuilder
    $stderr = New-Object System.Text.StringBuilder
    
    # Event handlers for output
    $scriptBlock = {
        if (-not [String]::IsNullOrEmpty($EventArgs.Data)) {
            $Event.MessageData.AppendLine($EventArgs.Data)
        }
    }
    
    $outEvent = Register-ObjectEvent -InputObject $process -EventName OutputDataReceived -Action $scriptBlock -MessageData $stdout
    $errEvent = Register-ObjectEvent -InputObject $process -EventName ErrorDataReceived -Action $scriptBlock -MessageData $stderr
    
    # Start process
    [void]$process.Start()
    $process.BeginOutputReadLine()
    $process.BeginErrorReadLine()
    
    # Wait for completion (max 2 hours)
    $timeout = 7200000 # 2 hours in milliseconds
    if (-not $process.WaitForExit($timeout)) {
        Write-Log "Process timeout after 2 hours - terminating" "ERROR"
        $process.Kill()
        exit 1
    }
    
    # Cleanup events
    Unregister-Event -SourceIdentifier $outEvent.Name
    Unregister-Event -SourceIdentifier $errEvent.Name
    
    $exitCode = $process.ExitCode
    $endTime = Get-Date
    $duration = ($endTime - $startTime).TotalSeconds
    
    Write-Log "Process completed in $([math]::Round($duration, 2)) seconds" "INFO"
    Write-Log "Exit code: $exitCode" "INFO"
    
    # Check exit code
    if ($exitCode -eq 0) {
        Write-Log "Maintenance completed successfully" "SUCCESS"
        
        # Create marker file to prevent re-run this month
        New-Item -Path $markerFile -ItemType File -Force | Out-Null
        $markerContent = @"
Maintenance completed successfully
Date: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
Exit Code: $exitCode
Duration: $([math]::Round($duration, 2)) seconds
"@
        Set-Content -Path $markerFile -Value $markerContent
        
        Write-Log "Marker file created: $markerFile" "SUCCESS"
        Write-Log "Task will not run again until next month" "INFO"
        
        # Cleanup old markers (keep last 6 months)
        $oldMarkers = Get-ChildItem $MarkerDir -Filter "*.completed" | 
            Where-Object { $_.Name -match '^\d{4}-\d{2}\.completed$' } |
            Sort-Object Name -Descending |
            Select-Object -Skip 6
        
        foreach ($oldMarker in $oldMarkers) {
            Remove-Item $oldMarker.FullName -Force
            Write-Log "Cleaned up old marker: $($oldMarker.Name)" "INFO"
        }
        
        exit 0
        
    } elseif ($exitCode -eq 1) {
        Write-Log "Maintenance completed but report submission failed" "WARNING"
        Write-Log "Task will retry tomorrow" "INFO"
        exit 1
        
    } else {
        Write-Log "Maintenance failed with exit code $exitCode" "ERROR"
        Write-Log "Task will retry tomorrow" "WARNING"
        exit 1
    }
    
} catch {
    Write-Log "Exception occurred: $($_.Exception.Message)" "ERROR"
    Write-Log "Task will retry tomorrow" "WARNING"
    exit 1
}
'@

# Write wrapper script
Set-Content -Path $wrapperScript -Value $wrapperContent -Encoding UTF8

Write-Host "[OK] Wrapper script created: $wrapperScript" -ForegroundColor Green

# Delete existing task if exists
Write-Host "`n[INFO] Checking for existing task..." -ForegroundColor Yellow
$existingTask = schtasks /Query /TN "$TaskName" 2>$null
if ($LASTEXITCODE -eq 0) {
    Write-Host "[INFO] Removing existing task..." -ForegroundColor Yellow
    schtasks /Delete /TN "$TaskName" /F | Out-Null
}

# Create XML for scheduled task (compatible with Windows 7+)
Write-Host "[CREATE] Creating scheduled task..." -ForegroundColor Yellow

# Use schtasks for maximum compatibility
# Create task that runs daily at specified time during first week
$taskCommand = "powershell.exe"
$taskArgs = "-NoProfile -ExecutionPolicy Bypass -File `"$wrapperScript`""

# Create task using schtasks (PS 2.0 compatible)
# Task runs daily, but wrapper script checks if it should actually execute

$result = schtasks /Create `
    /TN "$TaskName" `
    /TR "`"$taskCommand`" $taskArgs" `
    /SC DAILY `
    /ST $Time `
    /RU "SYSTEM" `
    /RL HIGHEST `
    /F

if ($LASTEXITCODE -eq 0) {
    Write-Host "`n[SUCCESS] Task Scheduler created successfully!" -ForegroundColor Green
    
    Write-Host "`nTask Details:" -ForegroundColor Cyan
    Write-Host "  Name: $TaskName" -ForegroundColor White
    Write-Host "  Schedule: Daily at $Time (auto-skip if already run this month)" -ForegroundColor White
    Write-Host "  Run Window: 1st-7th of each month" -ForegroundColor White
    Write-Host "  Run As: SYSTEM (Highest Privileges)" -ForegroundColor White
    Write-Host "  Wrapper: $wrapperScript" -ForegroundColor White
    Write-Host "  Markers: $markerDir" -ForegroundColor White
    
    Write-Host "`nHow it works:" -ForegroundColor Cyan
    Write-Host "  1. Task runs daily at $Time" -ForegroundColor Gray
    Write-Host "  2. Wrapper checks if current date is 1-7" -ForegroundColor Gray
    Write-Host "  3. Wrapper checks if already completed this month" -ForegroundColor Gray
    Write-Host "  4. Wrapper checks internet connectivity" -ForegroundColor Gray
    Write-Host "  5. If all OK, runs maintenance and creates marker" -ForegroundColor Gray
    Write-Host "  6. Will retry next day if failed or no internet" -ForegroundColor Gray
    Write-Host "  7. Skips execution if already completed this month" -ForegroundColor Gray
    
    # Show task info
    Write-Host "`nScheduled Task Query:" -ForegroundColor Cyan
    schtasks /Query /TN "$TaskName" /FO LIST /V | Select-String "Task To Run|Next Run Time|Status"
    
    Write-Host "`n"
    $runNow = Read-Host "Do you want to test the task now? (Y/N)"
    if ($runNow -eq "Y" -or $runNow -eq "y") {
        Write-Host "`n[TESTING] Starting task..." -ForegroundColor Yellow
        Write-Host "[INFO] This will check all conditions and run if valid" -ForegroundColor Gray
        Write-Host "[INFO] Check the log file for details:`n  $markerDir\maintenance.log`n" -ForegroundColor Gray
        
        schtasks /Run /TN "$TaskName"
        Start-Sleep -Seconds 3
        
        Write-Host "[INFO] Task started in background" -ForegroundColor Green
        Write-Host "[INFO] Check Task Scheduler or log file for results" -ForegroundColor Cyan
    }
    
} else {
    Write-Host "`n[ERROR] Failed to create scheduled task!" -ForegroundColor Red
    Write-Host "Error code: $LASTEXITCODE" -ForegroundColor Red
    pause
    exit 1
}

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Installation Complete!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "`nManage task via:" -ForegroundColor Cyan
Write-Host "  taskschd.msc" -ForegroundColor White
Write-Host "`nView logs:" -ForegroundColor Cyan
Write-Host "  $markerDir\maintenance.log" -ForegroundColor White
Write-Host "`nManual test:" -ForegroundColor Cyan
Write-Host "  schtasks /Run /TN `"$TaskName`"" -ForegroundColor White
pause
