# ========================================
# MONTHLY MAINTENANCE WRAPPER SCRIPT
# ========================================
# Compatible: PowerShell 2.0+
# Purpose: Run maintenance once per month with internet check and retry

param(
    [string]$ExePath = "C:\laragon\www\Scripts\MaintenanceApp\bin\Release\MaintenanceApp.exe",
    [string]$MarkerDir = "C:\laragon\www\Scripts\MaintenanceApp\TaskScheduler\Markers"
)

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
