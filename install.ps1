# install.ps1 - Install Whisper Auto-Transcriber as a Windows service

param(
    [string]$ServiceName = "WhisperTranscriber",
    [string]$DisplayName = "Whisper Auto-Transcriber",
    [string]$ScriptPath = "C:\whisper\whisper-service.ps1",
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

function Ensure-Directory {
    param([string]$Path)
    if (-not (Test-Path $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
        Write-Host "Created directory: $Path"
    }
}

# Ensure required directories
Ensure-Directory "C:\whisper"
Ensure-Directory "C:\whisper\transcripts"
Ensure-Directory "C:\whisper\completed"
Ensure-Directory "C:\whisper\logs"
Ensure-Directory "C:\whisper\service"

# Directory structure creation
$WatchPath = "C:\whisper\watch"
$OutputPath = "C:\whisper\output"
$CompletedPath = "C:\whisper\completed"
$LogPath = "C:\whisper\logs"

$requiredDirs = @($WatchPath, $OutputPath, $CompletedPath, $LogPath)
foreach ($dir in $requiredDirs) {
    try {
        if (-not (Test-Path $dir)) {
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
            Write-Host "Created directory: $dir" -ForegroundColor Green
        } else {
            Write-Host "Directory already exists: $dir" -ForegroundColor Cyan
        }
    } catch {
        Write-Host "ERROR: Failed to create directory: $dir. $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
}

# --- Robust Whisper CLI presence check (avoids Unicode and argument errors) ---
if (-not (Get-Command whisper -ErrorAction SilentlyContinue)) {
    Write-Host "ERROR: Whisper CLI is not available or not in PATH." -ForegroundColor Red
    Write-Host "Please install Whisper CLI and ensure it is accessible from the command line." -ForegroundColor Yellow
    exit 1
}

# --- Robust NSSM presence check ---
# --- Automatic NSSM Download and Install if Missing ---
$NSSMUrl = "https://nssm.cc/release/nssm-2.24.zip"
$NSSMZip = "$env:TEMP\nssm.zip"
$NSSMExtractDir = "$env:TEMP\nssm_extract"

if (-not $NSSMPath -or [string]::IsNullOrWhiteSpace($NSSMPath)) {
    $NSSMPath = "C:\whisper\nssm"
}

if (-not (Test-Path (Join-Path $NSSMPath 'nssm.exe'))) {
    Write-Host "NSSM not found. Downloading from $NSSMUrl..." -ForegroundColor Yellow
    Invoke-WebRequest -Uri $NSSMUrl -OutFile $NSSMZip
    Expand-Archive -Path $NSSMZip -DestinationPath $NSSMExtractDir -Force

    # Find the correct nssm.exe (prefer win64, fallback to win32)
    $nssmExe = Get-ChildItem -Path $NSSMExtractDir -Recurse -Filter nssm.exe | Where-Object { $_.FullName -match "win64|win32" } | Select-Object -First 1

    if ($nssmExe) {
        if (-not (Test-Path $NSSMPath)) { New-Item -ItemType Directory -Path $NSSMPath | Out-Null }
        Copy-Item $nssmExe.FullName -Destination (Join-Path $NSSMPath 'nssm.exe') -Force
        Write-Host "NSSM installed to $NSSMPath." -ForegroundColor Green
    } else {
        Write-Host "ERROR: Could not find nssm.exe in the downloaded archive." -ForegroundColor Red
        exit 1
    }
    # Cleanup
    Remove-Item $NSSMZip -Force
    Remove-Item $NSSMExtractDir -Recurse -Force
}

# --- Normalize NSSMPath to always be a directory ---
if ($NSSMPath -like '*nssm.exe') {
    $NSSMPath = Split-Path $NSSMPath -Parent
}

# --- Set NSSM executable path variable for robust invocation ---
$NSSMExePath = Join-Path $NSSMPath 'nssm.exe'
Write-Host "DEBUG: NSSMExePath is '$NSSMExePath'"
if (-not (Test-Path $NSSMExePath)) {
    Write-Host "ERROR: NSSM executable not found at $NSSMExePath" -ForegroundColor Red
    exit 1
}

# --- Automatic Python Installation if Missing ---
if (-not (Get-Command python -ErrorAction SilentlyContinue)) {
    Write-Host "Python not found. Downloading and installing Python..." -ForegroundColor Yellow
    $pythonInstallerUrl = "https://www.python.org/ftp/python/3.12.3/python-3.12.3-amd64.exe"
    $pythonInstaller = "$env:TEMP\python-installer.exe"
    Invoke-WebRequest -Uri $pythonInstallerUrl -OutFile $pythonInstaller
    Start-Process -FilePath $pythonInstaller -ArgumentList "/quiet InstallAllUsers=1 PrependPath=1 Include_test=0" -Wait
    Remove-Item $pythonInstaller -Force
    Write-Host "Python installed. You may need to restart your shell for PATH changes to take effect." -ForegroundColor Green
}

# --- Automatic pip Installation if Missing ---
if (-not (Get-Command pip -ErrorAction SilentlyContinue)) {
    Write-Host "pip not found. Installing pip..." -ForegroundColor Yellow
    try {
        python -m ensurepip
    } catch {
        $getPipUrl = "https://bootstrap.pypa.io/get-pip.py"
        $getPipScript = "$env:TEMP\get-pip.py"
        Invoke-WebRequest -Uri $getPipUrl -OutFile $getPipScript
        python $getPipScript
        Remove-Item $getPipScript -Force
    }
    Write-Host "pip installed." -ForegroundColor Green
}

# --- Automatic Whisper CLI Installation if Missing ---
if (-not (Get-Command whisper -ErrorAction SilentlyContinue)) {
    Write-Host "Whisper CLI not found. Installing via pip..." -ForegroundColor Yellow
    pip install git+https://github.com/openai/whisper.git
    Write-Host "Whisper CLI installed." -ForegroundColor Green
}

# --- Automatic ffmpeg Installation if Missing ---
if (-not (Get-Command ffmpeg -ErrorAction SilentlyContinue)) {
    Write-Host "ffmpeg not found. Downloading and installing ffmpeg..." -ForegroundColor Yellow
    $ffmpegUrl = "https://www.gyan.dev/ffmpeg/builds/ffmpeg-release-essentials.zip"
    $ffmpegZip = "$env:TEMP\ffmpeg.zip"
    $ffmpegExtractDir = "$env:TEMP\ffmpeg_extract"
    Invoke-WebRequest -Uri $ffmpegUrl -OutFile $ffmpegZip
    Expand-Archive -Path $ffmpegZip -DestinationPath $ffmpegExtractDir -Force
    $ffmpegBin = Get-ChildItem -Path $ffmpegExtractDir -Recurse -Filter ffmpeg.exe | Select-Object -First 1
    $ffmpegTarget = "C:\whisper\ffmpeg"
    if (-not (Test-Path $ffmpegTarget)) { New-Item -ItemType Directory -Path $ffmpegTarget | Out-Null }
    Copy-Item $ffmpegBin.FullName -Destination (Join-Path $ffmpegTarget 'ffmpeg.exe') -Force
    # Add to PATH for current session
    $env:PATH = "$ffmpegTarget;" + $env:PATH
    Write-Host "ffmpeg installed to $ffmpegTarget and added to PATH for this session." -ForegroundColor Green
    Remove-Item $ffmpegZip -Force
    Remove-Item $ffmpegExtractDir -Recurse -Force
}

# --- Determine script directory for robust file operations ---
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition

# --- Robust service script copy ---
$serviceScript = Join-Path $ScriptDir 'whisper-service.ps1'
$destScript = 'C:\whisper\service\whisper-service.ps1'
if (-not (Test-Path $serviceScript)) {
    Write-Host "ERROR: Could not find $serviceScript. Please ensure whisper-service.ps1 is in the same directory as install.ps1." -ForegroundColor Red
    exit 1
}
Copy-Item -Path $serviceScript -Destination $destScript -Force
Write-Host "Copied whisper-service.ps1 to $destScript" -ForegroundColor Green

# NSSM Service Configuration and Registration
$ServiceName = "WhisperTranscriber"
$DisplayName = "Whisper Auto-Transcriber Service"
$Description = "Automated audio transcription service using OpenAI Whisper CLI. Monitors a directory and transcribes new audio files."
$ServiceDir = "C:\whisper\service"
$AppDirectory = $ServiceDir
$AppParameters = "-ExecutionPolicy Bypass -File `"$destScript`""

try {
    # First, try to stop and remove any existing service
    Write-Host "Checking for existing service..." -ForegroundColor Cyan
    try {
        $existingService = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
        if ($existingService) {
            Write-Host "Stopping existing service..." -ForegroundColor Yellow
            Stop-Service -Name $ServiceName -Force -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 2
        }
    } catch {
        # Service doesn't exist, which is fine
    }
    
    # Remove service using NSSM
    Write-Host "Removing existing service configuration..." -ForegroundColor Yellow
    & "$NSSMExePath" remove $ServiceName confirm | Out-Null
    Start-Sleep -Seconds 3  # Wait for removal to complete
    
    # Install new service
    Write-Host "Installing new service..." -ForegroundColor Yellow
    & "$NSSMExePath" install $ServiceName powershell.exe $AppParameters | Out-Null
    Start-Sleep -Seconds 1
    
    # Configure service settings
    Write-Host "Configuring service settings..." -ForegroundColor Yellow
    & "$NSSMExePath" set $ServiceName AppDirectory $AppDirectory | Out-Null
    & "$NSSMExePath" set $ServiceName DisplayName "$DisplayName" | Out-Null
    & "$NSSMExePath" set $ServiceName Description "$Description" | Out-Null
    & "$NSSMExePath" set $ServiceName Start SERVICE_AUTO_START | Out-Null
    
    Write-Host "NSSM service '$ServiceName' installed and configured successfully." -ForegroundColor Green
} catch {
    Write-Host "ERROR: Failed to configure NSSM service: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
} 

# Installation Feedback and Summary
try {
    $service = Get-Service -Name $ServiceName -ErrorAction Stop
    if ($service.Status -eq 'Stopped') {
        Write-Host "Service '$ServiceName' is installed but not running." -ForegroundColor Yellow
        Write-Host "You can start the service with: nssm start $ServiceName" -ForegroundColor Cyan
    } elseif ($service.Status -eq 'Running') {
        Write-Host "Service '$ServiceName' is running." -ForegroundColor Green
    } else {
        Write-Host "Service '$ServiceName' status: $($service.Status)" -ForegroundColor Yellow
    }
    Write-Host "Installation complete!"
    Write-Host "Service Name: $ServiceName"
    Write-Host "Service Directory: $ServiceDir"
    Write-Host "Watch Directory: $WatchPath"
    Write-Host "Output Directory: $OutputPath"
    Write-Host "Completed Directory: $CompletedPath"
    Write-Host "Log Directory: $LogPath"
    Write-Host "To uninstall, use the provided uninstall.ps1 script or run: nssm remove $ServiceName confirm" -ForegroundColor Cyan
} catch {
    Write-Host "ERROR: Service '$ServiceName' was not found after installation. Please check the NSSM logs and try again." -ForegroundColor Red
    exit 1
} 