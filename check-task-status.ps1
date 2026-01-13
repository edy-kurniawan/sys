# ========================================
# CHECK TASK STATUS
# ========================================
# View status of monthly maintenance task

param(
    [string]$TaskName = "PC Maintenance Monthly Report"
)

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Monthly Maintenance Task Status" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

# Check if task exists
$existingTask = schtasks /Query /TN "$TaskName" 2>$null
if ($LASTEXITCODE -ne 0) {
    Write-Host "[ERROR] Task '$TaskName' not found!" -ForegroundColor Red
    Write-Host "Run install-monthly-task.ps1 to create it." -ForegroundColor Yellow
    pause
    exit 1
}

# Show task details
Write-Host "Task Details:" -ForegroundColor Cyan
schtasks /Query /TN "$TaskName" /FO LIST /V | Select-String "Task To Run|Next Run Time|Last Run Time|Last Result|Status|Run As User"

# Check marker directory
$markerDir = Join-Path $PSScriptRoot "TaskScheduler\Markers"
if (Test-Path $markerDir) {
    Write-Host "`nExecution History (Last 6 months):" -ForegroundColor Cyan
    
    $markers = Get-ChildItem $markerDir -Filter "*.completed" -ErrorAction SilentlyContinue |
        Sort-Object Name -Descending |
        Select-Object -First 6
    
    if ($markers) {
        foreach ($marker in $markers) {
            $month = $marker.BaseName
            $content = Get-Content $marker.FullName -Raw -ErrorAction SilentlyContinue
            
            Write-Host "`n  Month: $month" -ForegroundColor Green
            if ($content) {
                $content -split "`n" | ForEach-Object {
                    if ($_) {
                        Write-Host "    $_" -ForegroundColor Gray
                    }
                }
            }
        }
    } else {
        Write-Host "  No execution history found" -ForegroundColor Yellow
        Write-Host "  Task has not completed successfully yet" -ForegroundColor Gray
    }
    
    # Check current month status
    $currentMonth = Get-Date -Format "yyyy-MM"
    $currentMarker = Join-Path $markerDir "$currentMonth.completed"
    
    Write-Host "`nCurrent Month ($currentMonth):" -ForegroundColor Cyan
    if (Test-Path $currentMarker) {
        Write-Host "  Status: COMPLETED ✓" -ForegroundColor Green
        Write-Host "  Will not run again this month" -ForegroundColor Gray
    } else {
        Write-Host "  Status: PENDING" -ForegroundColor Yellow
        
        $currentDay = (Get-Date).Day
        if ($currentDay -ge 1 -and $currentDay -le 7) {
            Write-Host "  Current date: Day $currentDay (in valid window)" -ForegroundColor Green
            Write-Host "  Task will attempt to run today at scheduled time" -ForegroundColor Cyan
        } else {
            Write-Host "  Current date: Day $currentDay (outside window 1-7)" -ForegroundColor Gray
            Write-Host "  Next opportunity: 1st of next month" -ForegroundColor Yellow
        }
    }
    
    # Check log file
    $logFile = Join-Path $markerDir "maintenance.log"
    if (Test-Path $logFile) {
        Write-Host "`nRecent Log Entries (Last 20 lines):" -ForegroundColor Cyan
        Get-Content $logFile -Tail 20 -ErrorAction SilentlyContinue | ForEach-Object {
            Write-Host "  $_" -ForegroundColor Gray
        }
        
        Write-Host "`nFull log: $logFile" -ForegroundColor Gray
    }
    
} else {
    Write-Host "`nNo execution history found" -ForegroundColor Yellow
}

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Commands:" -ForegroundColor Cyan
Write-Host "  Run now (test):  schtasks /Run /TN `"$TaskName`"" -ForegroundColor White
Write-Host "  View in GUI:     taskschd.msc" -ForegroundColor White
Write-Host "  Uninstall:       .\uninstall-monthly-task.ps1" -ForegroundColor White
Write-Host "========================================" -ForegroundColor Cyan
pause
