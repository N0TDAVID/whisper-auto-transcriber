# uninstall.ps1 - Uninstall Whisper Auto-Transcriber Windows service

param(
    [string]$ServiceName = "WhisperTranscriber",
    [string]$NSSMPath = "C:\nssm\nssm.exe",
    [string]$InstallPath = "C:\whisper",
    [switch]$RemoveDirs,
    [switch]$RemovePythonPackages,
    [switch]$RemoveNSSM
)

# --- Administrator Privilege Validation ---
$windowsIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$windowsPrincipal = New-Object System.Security.Principal.WindowsPrincipal($windowsIdentity)
$adminRole = [System.Security.Principal.WindowsBuiltInRole]::Administrator
if (-not $windowsPrincipal.IsInRole($adminRole)) {
    Write-Host "ERROR: This script must be run as Administrator." -ForegroundColor Red
    exit 1
}

# Check NSSM (only required if service exists)
$nssmExists = Test-Path $NSSMPath
if (-not $nssmExists) {
    Write-Host "NSSM not found at $NSSMPath. Service removal may be skipped." -ForegroundColor Yellow
}

# --- Service Status Check and Stopping ---
if ($nssmExists) {
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
    }
} else {
    Write-Host "Skipping service removal (NSSM not found)." -ForegroundColor Yellow
}

# --- Remove Python Packages ---
if ($RemovePythonPackages) {
    Write-Host "`n=== Removing Python Packages ===" -ForegroundColor Cyan
    Write-Host "WARNING: This will uninstall WhisperX, PyTorch, and cuDNN packages from your Python environment." -ForegroundColor Yellow
    $confirm = Read-Host "Type 'YES' to confirm Python package removal, or press Enter to skip"
    if ($confirm -eq 'YES') {
        try {
            Write-Host "Uninstalling WhisperX..." -ForegroundColor Cyan
            pip uninstall whisperx -y 2>&1 | Out-Null
            Write-Host "WhisperX uninstalled." -ForegroundColor Green
            
            Write-Host "Uninstalling cuDNN libraries..." -ForegroundColor Cyan
            pip uninstall nvidia-cudnn-cu11 -y 2>&1 | Out-Null
            Write-Host "cuDNN libraries uninstalled." -ForegroundColor Green
            
            Write-Host "Uninstalling PyTorch..." -ForegroundColor Cyan
            pip uninstall torch torchvision torchaudio -y 2>&1 | Out-Null
            Write-Host "PyTorch uninstalled." -ForegroundColor Green
            
            Write-Host "Python packages successfully removed." -ForegroundColor Green
        } catch {
            Write-Host "Error removing Python packages: $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "You may need to manually uninstall packages using: pip uninstall <package-name>" -ForegroundColor Yellow
        }
    } else {
        Write-Host "Skipped Python package removal." -ForegroundColor Cyan
    }
}

# --- Remove Directories ---
if ($RemoveDirs) {
    Write-Host "`n=== Removing Directories ===" -ForegroundColor Cyan
    $dirs = @(
        @{Path = "$InstallPath\watch"; Label = "Watch Directory"},
        @{Path = "$InstallPath\transcripts"; Label = "Transcripts Directory"},
        @{Path = "$InstallPath\completed"; Label = "Completed Directory"},
        @{Path = "$InstallPath\failed"; Label = "Failed Directory"},
        @{Path = "$InstallPath\logs"; Label = "Logs Directory"},
        @{Path = "$InstallPath\service"; Label = "Service Directory"},
        @{Path = $InstallPath; Label = "Root Service Directory"}
    )
    
    # Remove root directory last
    $otherDirs = $dirs | Where-Object { $_.Path -ne $InstallPath }
    
    foreach ($dirInfo in $otherDirs) {
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
                    Write-Host "Failed to remove directory $dir`: $($_.Exception.Message)" -ForegroundColor Red
                }
            } else {
                Write-Host "Skipped removal of $label ($dir)." -ForegroundColor Cyan
            }
        } else {
            Write-Host "$label ($dir) does not exist. Skipping." -ForegroundColor Gray
        }
    }
    
    # Remove root directory last (if all subdirectories are removed)
    if (Test-Path $InstallPath) {
        $remainingDirs = Get-ChildItem -Path $InstallPath -Directory -ErrorAction SilentlyContinue
        if ($remainingDirs.Count -eq 0) {
            Write-Host "`nWARNING: Root directory ($InstallPath) will be permanently deleted. This cannot be undone." -ForegroundColor Yellow
            $confirm = Read-Host "Type 'YES' to confirm deletion of root directory, or press Enter to skip"
            if ($confirm -eq 'YES') {
                try {
                    Remove-Item $InstallPath -Recurse -Force
                    Write-Host "Removed root directory: $InstallPath" -ForegroundColor Green
                } catch {
                    Write-Host "Failed to remove root directory $InstallPath`: $($_.Exception.Message)" -ForegroundColor Red
                }
            } else {
                Write-Host "Skipped removal of root directory ($InstallPath)." -ForegroundColor Cyan
            }
        } else {
            Write-Host "`nRoot directory ($InstallPath) contains remaining subdirectories. Skipping removal." -ForegroundColor Yellow
        }
    }
}

# --- Remove NSSM (optional) ---
if ($RemoveNSSM -and $nssmExists) {
    Write-Host "`n=== Removing NSSM ===" -ForegroundColor Cyan
    Write-Host "WARNING: This will remove NSSM from $NSSMPath. This may affect other services using NSSM." -ForegroundColor Yellow
    $confirm = Read-Host "Type 'YES' to confirm NSSM removal, or press Enter to skip"
    if ($confirm -eq 'YES') {
        try {
            $nssmDir = Split-Path $NSSMPath -Parent
            Remove-Item $nssmDir -Recurse -Force
            Write-Host "NSSM removed from $nssmDir" -ForegroundColor Green
        } catch {
            Write-Host "Failed to remove NSSM: $($_.Exception.Message)" -ForegroundColor Red
        }
    } else {
        Write-Host "Skipped NSSM removal." -ForegroundColor Cyan
    }
}

# --- Summary ---
Write-Host "`n=== Uninstall Summary ===" -ForegroundColor Cyan
Write-Host "Service removal: Completed" -ForegroundColor Green
if ($RemovePythonPackages) {
    Write-Host "Python packages: Removed (if confirmed)" -ForegroundColor Green
} else {
    Write-Host "Python packages: Not removed (use -RemovePythonPackages to remove)" -ForegroundColor Yellow
}
if ($RemoveDirs) {
    Write-Host "Directories: Removed (if confirmed)" -ForegroundColor Green
} else {
    Write-Host "Directories: Not removed (use -RemoveDirs to remove)" -ForegroundColor Yellow
}
if ($RemoveNSSM) {
    Write-Host "NSSM: Removed (if confirmed)" -ForegroundColor Green
} else {
    Write-Host "NSSM: Not removed (use -RemoveNSSM to remove)" -ForegroundColor Yellow
}
Write-Host "`nUninstall process completed." -ForegroundColor Green 
