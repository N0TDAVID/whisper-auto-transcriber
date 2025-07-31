# whisper-service.ps1
# Whisper Auto-Transcriber Windows Service
# 
# Supported Audio Formats:
# - M4A (AAC)
# - MP3
# - WAV
# - FLAC
# - AAC
# - OGG
# - WMA
# - M4B (Audiobook)
# - WebM
#
# =============================
# Configuration Parameters
# =============================
[string]$WatchPath     = "C:\whisper\watch"
[string]$OutputPath    = "C:\whisper\transcripts"
[string]$CompletedPath = "C:\whisper\completed"
[string]$FailedPath    = "C:\whisper\failed"
[string]$LogPath       = "C:\whisper\logs"
[string]$Language      = "en"
[int]$CheckInterval    = 5  # seconds

# =============================
# Validation Functions
# =============================
function Test-Directory {
    param([string]$Path)
    if ([string]::IsNullOrEmpty($Path)) {
        Write-Error "Path parameter is null or empty"
        exit 1
    }
    if (-not (Test-Path $Path)) {
        try {
            New-Item -ItemType Directory -Path $Path -Force | Out-Null
            Write-Host "Created missing directory: $Path"
        } catch {
            Write-Error "Failed to create directory: $Path. $_"
            exit 1
        }
    }
}

function Validate-LanguageCode {
    param([string]$Lang)
    if ($Lang -notmatch '^[a-z]{2}$') {
        Write-Error "Invalid language code: $Lang. Must be a two-letter code (e.g., 'en')."
        exit 1
    }
}

function Validate-CheckInterval {
    param([int]$Interval)
    if ($Interval -le 0) {
        Write-Error "CheckInterval must be a positive integer."
        exit 1
    }
}

# =============================
# Run Validations
# =============================
Test-Directory $WatchPath
Test-Directory $OutputPath
Test-Directory $CompletedPath
Test-Directory $FailedPath
Test-Directory $LogPath
Validate-LanguageCode $Language
Validate-CheckInterval $CheckInterval

Write-Host "Configuration parameters initialized and validated successfully."

# =============================
# Logging Function
# =============================
function Write-Log {
    param(
        [Parameter(Mandatory=$true)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR')][string]$Level = 'INFO'
    )
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logLine = "[$timestamp] [$Level] $Message"
    $logFileName = "whisper-service-" + (Get-Date -Format 'yyyyMMdd') + ".log"
    $logFile = Join-Path $LogPath $logFileName
    
    try {
        # Ensure log directory exists
        if (-not (Test-Path $LogPath)) {
            New-Item -ItemType Directory -Path $LogPath -Force | Out-Null
        }
        # Write to log file
        Add-Content -Path $logFile -Value $logLine -ErrorAction SilentlyContinue
    } catch {
        Write-Host "[ERROR][Log] Failed to write to log file: $logFile. $($_.Exception.Message)" -ForegroundColor Red
    }
    
    # Console output with color coding
    switch ($Level) {
        'INFO'  { Write-Host $logLine -ForegroundColor Green }
        'WARN'  { Write-Host $logLine -ForegroundColor Yellow }
        'ERROR' { Write-Host $logLine -ForegroundColor Red }
        default { Write-Host $logLine }
    }
}

# =============================
# Utility Functions
# =============================
function Get-UniqueFileName {
    param(
        [string]$Directory,
        [string]$BaseName,
        [string]$Extension
    )
    $counter = 1
    $fileName = $BaseName + $Extension
    $filePath = Join-Path $Directory $fileName
    
    while (Test-Path $filePath) {
        $fileName = "${BaseName}_${counter}${Extension}"
        $filePath = Join-Path $Directory $fileName
        $counter++
    }
    
    return $filePath
}

# =============================
# File Processing Functions
# =============================
function Process-AudioFile {
    param(
        [Parameter(Mandatory=$true)][string]$FilePath
    )
    
    # Validate input file
    if (-not (Test-Path $FilePath)) {
        Write-Log -Message "Audio file does not exist: $FilePath" -Level "ERROR"
        return
    }
    
    # Validate supported audio formats
    $supportedExtensions = @('.m4a', '.wav', '.mp3', '.flac', '.aac', '.ogg', '.wma', '.m4b', '.webm')
    $fileExt = [System.IO.Path]::GetExtension($FilePath)
    if ($null -ne $fileExt) { $fileExt = $fileExt.ToLower() } else { $fileExt = '' }
    if ($supportedExtensions -notcontains $fileExt) {
        Write-Log -Message "Unsupported audio format: $fileExt. Supported: $($supportedExtensions -join ', ')" -Level "ERROR"
        return
    }
    
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
    Write-Log -Message "Starting transcription for: $FilePath" -Level "INFO"
    
    # Check if whisper CLI is available
    $whisperCmd = Get-Command whisper -ErrorAction SilentlyContinue
    if (-not $whisperCmd) {
        Write-Log -Message "Whisper CLI executable not found in PATH. Please install or add to PATH." -Level "ERROR"
        return
    }
    
    try {
        # Build whisper command arguments with quoted filename
        $fileName = [System.IO.Path]::GetFileName($FilePath)
        $directory = [System.IO.Path]::GetDirectoryName($FilePath)
        $quotedFilePath = Join-Path $directory "`"$fileName`""
        
        $whisperArgs = @(
            "--model", "medium",
            "--language", $Language,
            "--output_format", "txt",
            "--output_dir", $OutputPath,
            $quotedFilePath
        )
        
        Write-Log -Message "Running whisper command: $($whisperCmd.Source) $($whisperArgs -join ' ')" -Level "INFO"
        
        # Execute whisper command directly using PowerShell
        $process = Start-Process -FilePath $whisperCmd.Source -ArgumentList $whisperArgs -PassThru -NoNewWindow -RedirectStandardOutput "$OutputPath\temp_output.txt" -RedirectStandardError "$OutputPath\temp_error.txt"
        
        # Wait for process with timeout (10 minutes)
        $timeout = 600
        $waited = 0
        while (-not $process.HasExited -and $waited -lt $timeout) {
            Start-Sleep -Seconds 2
            $waited += 2
        }
        
        if (-not $process.HasExited) {
            $process.Kill()
            Write-Log -Message "Whisper process timed out after $timeout seconds for $FilePath" -Level "ERROR"
            throw "Process timeout"
        }
        
        if ($process.ExitCode -eq 0) {
            Write-Log -Message "Successfully processed: $FilePath" -Level "INFO"
            
            # Check temp output and error files before cleanup
            if (Test-Path "$OutputPath\temp_output.txt") {
                $outputContent = Get-Content "$OutputPath\temp_output.txt" -ErrorAction SilentlyContinue
                Write-Log -Message "Whisper output: $($outputContent -join ' ')" -Level "INFO"
            }
            if (Test-Path "$OutputPath\temp_error.txt") {
                $errorContent = Get-Content "$OutputPath\temp_error.txt" -ErrorAction SilentlyContinue
                Write-Log -Message "Whisper errors: $($errorContent -join ' ')" -Level "INFO"
            }
            
            # Check if transcript file was actually created
            $expectedTranscriptFile = Join-Path $OutputPath "$baseName.txt"
            if (Test-Path $expectedTranscriptFile) {
                Write-Log -Message "Transcript file created: $expectedTranscriptFile" -Level "INFO"
            } else {
                Write-Log -Message "WARNING: Expected transcript file not found: $expectedTranscriptFile" -Level "WARN"
                # Check what files were actually created in the output directory
                $createdFiles = Get-ChildItem -Path $OutputPath -Filter "*.txt" | Where-Object {$_.LastWriteTime -gt (Get-Date).AddMinutes(-5)}
                if ($createdFiles) {
                    Write-Log -Message "Recently created files in output directory: $($createdFiles.Name -join ', ')" -Level "INFO"
                } else {
                    Write-Log -Message "No recent transcript files found in output directory" -Level "WARN"
                }
            }
            
            # Move processed file to completed directory
            Move-Item -Path $FilePath -Destination (Join-Path $CompletedPath ([System.IO.Path]::GetFileName($FilePath))) -Force
            Write-Log -Message "Moved to completed: $FilePath" -Level "INFO"
        } else {
            $errorOutput = Get-Content "$OutputPath\temp_error.txt" -ErrorAction SilentlyContinue
            Write-Log -Message "Whisper CLI failed for ${FilePath} with exit code: $($process.ExitCode). Error: $errorOutput" -Level "ERROR"
            
            # Move failed file to failed directory
            Move-Item -Path $FilePath -Destination (Join-Path $FailedPath ([System.IO.Path]::GetFileName($FilePath))) -Force
            Write-Log -Message "Moved to failed: $FilePath" -Level "ERROR"
        }
        
        # Clean up temp files
        Remove-Item "$OutputPath\temp_output.txt" -ErrorAction SilentlyContinue
        Remove-Item "$OutputPath\temp_error.txt" -ErrorAction SilentlyContinue
        
    } catch {
        Write-Log -Message "Exception in Process-AudioFile for ${FilePath}: $($_.Exception.Message)" -Level "ERROR"
        
        # Move failed file to failed directory
        Move-Item -Path $FilePath -Destination (Join-Path $FailedPath ([System.IO.Path]::GetFileName($FilePath))) -Force
        Write-Log -Message "Moved to failed: $FilePath" -Level "ERROR"
    }
}

# =============================
# File System Watcher Functions
# =============================
function Start-FileSystemWatcher {
    param([string]$Path)
    
    if ($global:FileSystemWatcher) {
        $global:FileSystemWatcher.EnableRaisingEvents = $false
        $global:FileSystemWatcher.Dispose()
    }
    
    $global:FileSystemWatcher = New-Object System.IO.FileSystemWatcher
    $global:FileSystemWatcher.Path = $Path
    $global:FileSystemWatcher.Filter = "*.m4a;*.mp3;*.wav;*.flac;*.aac;*.ogg;*.wma;*.m4b;*.webm"
    $global:FileSystemWatcher.IncludeSubdirectories = $false
    $global:FileSystemWatcher.NotifyFilter = [System.IO.NotifyFilters]::FileName
    
    # Register event handlers
    Register-ObjectEvent -InputObject $global:FileSystemWatcher -EventName Created -Action {
        $filePath = $Event.SourceEventArgs.FullPath
        Write-Log -Message "New audio file detected: $filePath" -Level "INFO"
        # File will be processed in the main loop
    }
    
    $global:FileSystemWatcher.EnableRaisingEvents = $true
    Write-Log -Message "FileSystemWatcher started for path: $Path" -Level "INFO"
}

function Stop-FileSystemWatcher {
    if ($global:FileSystemWatcher) {
        try {
            $global:FileSystemWatcher.EnableRaisingEvents = $false
            $global:FileSystemWatcher.Dispose()
            Write-Log -Message "FileSystemWatcher stopped." -Level "INFO"
        } catch {
            Write-Log -Message "Error stopping FileSystemWatcher: $($_.Exception.Message)" -Level "ERROR"
        }
    }
}

# =============================
# Graceful Shutdown Functions
# =============================
function Stop-ServiceGracefully {
    Write-Log -Message "Initiating graceful shutdown..." -Level "INFO"
    
    try {
        # Stop FileSystemWatcher
        Stop-FileSystemWatcher
        
        Write-Log -Message "Graceful shutdown complete." -Level "INFO"
    } catch {
        Write-Log -Message "Exception during graceful shutdown: $($_.Exception.Message)" -Level "ERROR"
    }
}

# Register graceful shutdown handler
$null = Register-EngineEvent PowerShell.Exiting -Action { Stop-ServiceGracefully }

# =============================
# Start the Service
# =============================
try {
    Write-Log -Message "Starting Whisper Service..." -Level "INFO"
    
    # Start FileSystemWatcher
    Start-FileSystemWatcher -Path $WatchPath
    
    Write-Log -Message "Service started successfully. Entering main loop..." -Level "INFO"
    
    # Main service loop
    while ($true) {
        try {
            # Check if FileSystemWatcher is still working
            if ($global:FileSystemWatcher -and -not $global:FileSystemWatcher.EnableRaisingEvents) {
                Write-Log -Message "FileSystemWatcher stopped, restarting..." -Level "WARN"
                Start-FileSystemWatcher -Path $WatchPath
            }
            
            # Process any audio files in the watch directory
            $audioExtensions = @("*.m4a", "*.mp3", "*.wav", "*.flac", "*.aac", "*.ogg", "*.wma", "*.m4b", "*.webm")
            $files = @()
            foreach ($ext in $audioExtensions) {
                $files += Get-ChildItem -Path $WatchPath -Filter $ext -ErrorAction SilentlyContinue
            }
            
            foreach ($file in $files) {
                Write-Log -Message "Processing file: $($file.FullName)" -Level "INFO"
                Process-AudioFile -FilePath $file.FullName
            }
            
            # Sleep for a short interval to prevent high CPU usage
            Start-Sleep -Seconds $CheckInterval
        } catch {
            Write-Log -Message "Error in main loop: $($_.Exception.Message)" -Level "ERROR"
            Start-Sleep -Seconds $CheckInterval
        }
    }
} catch {
    Write-Log -Message "Critical error in service: $($_.Exception.Message)" -Level "ERROR"
    throw
}
