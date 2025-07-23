# uninstall.ps1 - Uninstall Whisper Auto-Transcriber Windows service

param(
    [string]$ServiceName = "WhisperTranscriber",
    [string]$NSSMPath = "C:\nssm\nssm.exe",
    [switch]$RemoveDirs
)

# --- Administrator Privilege Validation ---
$windowsIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$windowsPrincipal = New-Object System.Security.Principal.WindowsPrincipal($windowsIdentity)
$adminRole = [System.Security.Principal.WindowsBuiltInRole]::Administrator
if (-not $windowsPrincipal.IsInRole($adminRole)) {
    Write-Host "ERROR: This script must be run as Administrator." -ForegroundColor Red
    exit 1
}

# Check NSSM
if (-not (Test-Path $NSSMPath)) {
    Write-Host "NSSM not found at $NSSMPath. Please download NSSM and place it there." -ForegroundColor Red
    exit 1
}

# --- Service Status Check and Stopping ---
try {
    $service = Get-Service -Name $ServiceName -ErrorAction Stop
    if ($service.Status -eq 'Stopped') {
        Write-Host "Service '$ServiceName' is already stopped." -ForegroundColor Yellow
    } else {
        Write-Host "Stopping service '$ServiceName'..." -ForegroundColor Cyan
        & $NSSMPath stop $ServiceName
        Write-Host "Service '$ServiceName' stopped." -ForegroundColor Green
    }
} catch {
    Write-Host "Service '$ServiceName' does not exist or is not installed." -ForegroundColor Yellow
}

# --- Remove Service ---
try {
    & $NSSMPath remove $ServiceName confirm
    Start-Sleep -Seconds 2 # Give Windows time to update service list
    try {
        $serviceCheck = Get-Service -Name $ServiceName -ErrorAction Stop
        Write-Host "WARNING: Service $ServiceName still exists after removal attempt." -ForegroundColor Yellow
    } catch {
        Write-Host "Service $ServiceName successfully removed from Windows Services." -ForegroundColor Green
    }
} catch {
    Write-Host "Failed to remove service: $($_.Exception.Message)" -ForegroundColor Red
    if (-not (Test-Path $NSSMPath)) {
        Write-Host "NSSM not found at $NSSMPath. Please ensure NSSM is installed and accessible." -ForegroundColor Red
    }
}

# Optionally remove directories with user confirmation
if ($RemoveDirs) {
    $dirs = @(
        @{Path = "C:\whisper\transcripts"; Label = "Transcripts Directory"},
        @{Path = "C:\whisper\completed"; Label = "Completed Directory"},
        @{Path = "C:\whisper\logs"; Label = "Logs Directory"},
        @{Path = "C:\whisper"; Label = "Root Service Directory"}
    )
    foreach ($dirInfo in $dirs) {
        $dir = $dirInfo.Path
        $label = $dirInfo.Label
        if (Test-Path $dir) {
            Write-Host "WARNING: $label ($dir) will be permanently deleted. This cannot be undone." -ForegroundColor Yellow
            $confirm = Read-Host "Type 'YES' to confirm deletion of $label, or press Enter to skip"
            if ($confirm -eq 'YES') {
                try {
                    Remove-Item $dir -Recurse -Force
                    Write-Host "Removed directory: $dir" -ForegroundColor Green
                } catch {
                    Write-Host "Failed to remove directory $dir: $($_.Exception.Message)" -ForegroundColor Red
                }
            } else {
                Write-Host "Skipped removal of $label ($dir)." -ForegroundColor Cyan
            }
        } else {
            Write-Host "$label ($dir) does not exist. Skipping." -ForegroundColor Gray
        }
    }
} 