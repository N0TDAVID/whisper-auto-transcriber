# cleanup-service.ps1 - Clean up stuck WhisperTranscriber service

param(
    [string]$ServiceName = "WhisperTranscriber",
    [string]$NSSMPath = "C:\whisper\nssm"
)

# Administrator privilege check
$windowsIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$windowsPrincipal = New-Object System.Security.Principal.WindowsPrincipal($windowsIdentity)
$isAdmin = $windowsPrincipal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "ERROR: This script must be run as an administrator." -ForegroundColor Red
    Write-Host "Right-click PowerShell and select 'Run as administrator', then re-run this script." -ForegroundColor Yellow
    exit 1
}

Write-Host "Cleaning up WhisperTranscriber service..." -ForegroundColor Cyan

# Set NSSM executable path
$NSSMExePath = Join-Path $NSSMPath 'nssm.exe'

# Step 1: Stop the service if it's running
try {
    $service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    if ($service) {
        Write-Host "Found service '$ServiceName'. Stopping it..." -ForegroundColor Yellow
        Stop-Service -Name $ServiceName -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 3
    } else {
        Write-Host "Service '$ServiceName' not found in Windows Services." -ForegroundColor Yellow
    }
} catch {
    Write-Host "Error checking service: $($_.Exception.Message)" -ForegroundColor Red
}

# Step 2: Remove service using NSSM
if (Test-Path $NSSMExePath) {
    Write-Host "Removing service using NSSM..." -ForegroundColor Yellow
    try {
        & "$NSSMExePath" remove $ServiceName confirm
        Start-Sleep -Seconds 3
        Write-Host "Service removed successfully." -ForegroundColor Green
    } catch {
        Write-Host "Error removing service with NSSM: $($_.Exception.Message)" -ForegroundColor Red
    }
} else {
    Write-Host "NSSM not found at $NSSMExePath" -ForegroundColor Red
}

# Step 3: Force remove from registry if needed
Write-Host "Cleaning up registry entries..." -ForegroundColor Yellow
try {
    $registryPath = "HKLM:\SYSTEM\CurrentControlSet\Services\$ServiceName"
    if (Test-Path $registryPath) {
        Remove-Item -Path $registryPath -Recurse -Force
        Write-Host "Registry entries removed." -ForegroundColor Green
    } else {
        Write-Host "No registry entries found." -ForegroundColor Yellow
    }
} catch {
    Write-Host "Error cleaning registry: $($_.Exception.Message)" -ForegroundColor Red
}

# Step 4: Verify cleanup
Start-Sleep -Seconds 2
try {
    $serviceCheck = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    if ($serviceCheck) {
        Write-Host "WARNING: Service still exists after cleanup." -ForegroundColor Red
        Write-Host "You may need to restart your computer to complete the cleanup." -ForegroundColor Yellow
    } else {
        Write-Host "Service cleanup completed successfully!" -ForegroundColor Green
        Write-Host "You can now run install.ps1 again." -ForegroundColor Cyan
    }
} catch {
    Write-Host "Service cleanup completed successfully!" -ForegroundColor Green
    Write-Host "You can now run install.ps1 again." -ForegroundColor Cyan
} 