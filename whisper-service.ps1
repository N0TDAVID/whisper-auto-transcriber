# whisper-service.ps1

# =============================
# Configuration Parameters
# =============================
[string]$WatchPath     = "C:\whisper"
[string]$OutputPath    = "C:\whisper\transcripts"
[string]$CompletedPath = "C:\whisper\completed"
[string]$LogPath       = "C:\whisper\logs"
[string]$Language      = "en"
[int]$CheckInterval    = 5  # seconds

# --- Enhancement: Add $FailedPath global variable for failed files ---
if (-not (Test-Path $CompletedPath)) {
    $global:CompletedPath = Join-Path $OutputPath 'completed'
    Ensure-Directory $global:CompletedPath
}
if (-not (Test-Path $FailedPath)) {
    $global:FailedPath = Join-Path $OutputPath 'failed'
    Ensure-Directory $global:FailedPath
}

# =============================
# Validation Functions
# =============================
function Test-Directory {
    param([string]$Path)
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
        [ValidateSet('INFO','WARN','ERROR')][string]$Level = 'INFO',
        [string]$LogPathOverride
    )
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logLine = "[$timestamp] [$Level] $Message"
    $logDir = if ($LogPathOverride) { Split-Path $LogPathOverride -Parent } else { $LogPath }
    $logFileName = "whisper-service-" + (Get-Date -Format 'yyyyMMdd') + ".log"
    $logFile = if ($LogPathOverride) { $LogPathOverride } else { Join-Path $logDir $logFileName }
    try {
        # Ensure log directory exists
        if (-not (Test-Path $logDir)) {
            New-Item -ItemType Directory -Path $logDir -Force | Out-Null
        }
        # Write to log file with file lock protection
        $fileStream = [System.IO.File]::Open($logFile, [System.IO.FileMode]::Append, [System.IO.FileAccess]::Write, [System.IO.FileShare]::ReadWrite)
        $writer = New-Object System.IO.StreamWriter($fileStream)
        $writer.WriteLine($logLine)
        $writer.Close()
        $fileStream.Close()
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

# Background job to clean up log files older than 7 days
Start-Job -ScriptBlock {
    $logDir = "$using:LogPath"
    while ($true) {
        $now = Get-Date
        Get-ChildItem -Path $logDir -Filter 'whisper-service-*.log' | Where-Object {
            ($now - $_.LastWriteTime).TotalDays -gt 7
        } | ForEach-Object {
            try {
                Remove-Item $_.FullName -Force
            } catch {}
        }
        Start-Sleep -Seconds 3600  # Run cleanup every hour
    }
} | Out-Null

# Example usage:
Write-Log -Message "Service script started." -Level "INFO"

# =============================
# Error Handling Framework
# =============================
function Write-ErrorLog {
    param(
        [Parameter(Mandatory=$true)][string]$Message,
        [string]$Category = 'General',
        [System.Management.Automation.ErrorRecord]$ErrorRecord
    )
    $stack = if ($ErrorRecord) { $ErrorRecord.ScriptStackTrace } else { $null }
    $fullMsg = if ($stack) { "$Message`nStackTrace: $stack" } else { $Message }
    Write-Log -Message "[$Category] $fullMsg" -Level "ERROR"
}

# =============================
# Main Service Loop
# =============================
$ServiceRunning = $true

# Trap Ctrl+C and termination signals for graceful shutdown
$null = Register-EngineEvent PowerShell.Exiting -Action {
    $global:ServiceRunning = $false
    Write-Log -Message "Service shutdown signal received. Cleaning up..." -Level "WARN"
}

try {
    Write-Log -Message "Entering main service loop." -Level "INFO"
    while ($ServiceRunning) {
        try {
            # Heartbeat log every interval
            Write-Log -Message "Service heartbeat. Monitoring $WatchPath for new files..." -Level "INFO"
            # (File monitoring and processing logic will be added in later tasks)
        } catch {
            Write-ErrorLog -Message "Error in service loop: $($_.Exception.Message)" -Category "ServiceLoop" -ErrorRecord $_
        }
        Start-Sleep -Seconds $CheckInterval
    }
    Write-Log -Message "Service loop exited. Performing final cleanup." -Level "WARN"
} catch {
    Write-ErrorLog -Message "Critical error in main service: $($_.Exception.Message)" -Category "Critical" -ErrorRecord $_
    exit 1
} finally {
    Write-Log -Message "Service script terminated." -Level "INFO"
} 

# =============================
# File System Monitoring (FileSystemWatcher)
# =============================
$global:Watcher = $null
$global:WatcherEvent = $null
$global:LastDetectedFiles = @{}
$global:WatcherHealthCheckInterval = 60  # seconds

function Start-FileSystemWatcher {
    param([string]$Path, [string]$Filter = '*.m4a')
    # Validate watch directory exists and is accessible
    if (-not (Test-Path $Path)) {
        Write-ErrorLog -Message "WatchPath does not exist: $Path" -Category "FileSystemWatcher"
        throw "WatchPath does not exist: $Path"
    }
    try {
        $acl = Get-Acl $Path
    } catch {
        Write-ErrorLog -Message "Insufficient permissions to access WatchPath: $Path" -Category "FileSystemWatcher" -ErrorRecord $_
        throw "Insufficient permissions to access WatchPath: $Path"
    }
    if (${global:Watcher}) { ${global:Watcher}.Dispose() }
    ${global:Watcher} = New-Object System.IO.FileSystemWatcher $Path, $Filter
    ${global:Watcher}.IncludeSubdirectories = $false
    ${global:Watcher}.EnableRaisingEvents = $true
    Write-Log -Message "FileSystemWatcher started on $Path for $Filter files." -Level "INFO"

    # Register event handler with unique identifier
    if (${global:WatcherEvent}) { Unregister-Event -SourceIdentifier FileCreated -ErrorAction SilentlyContinue }
    ${global:WatcherEvent} = Register-ObjectEvent -InputObject ${global:Watcher} -EventName Created -SourceIdentifier FileCreated -Action {
        param($sender, $eventArgs)
        $filePath = $eventArgs.FullPath
        # File existence and size validation
        if (-not (Test-Path $filePath)) {
            Write-Log -Message "Detected file does not exist: $filePath" -Level "WARN"
            return
        }
        $fileInfo = Get-Item $filePath
        if ($fileInfo.Length -eq 0) {
            Write-Log -Message "Detected file is empty: $filePath" -Level "WARN"
            return
        }
        # Duplicate detection (avoid re-queuing same file rapidly)
        $now = Get-Date
        if ($global:LastDetectedFiles.ContainsKey($filePath)) {
            $lastTime = $global:LastDetectedFiles[$filePath]
            if (($now - $lastTime).TotalSeconds -lt 10) {
                Write-Log -Message "Duplicate detection event ignored for: $filePath" -Level "WARN"
                return
            }
        }
        $global:LastDetectedFiles[$filePath] = $now
        Write-Log -Message "Detected new file: $filePath" -Level "INFO"
        Add-ToQueue $filePath
    }
}

function Stop-FileSystemWatcher {
    try {
        if (${global:WatcherEvent}) {
            Unregister-Event -SourceIdentifier FileCreated -ErrorAction SilentlyContinue
            ${global:WatcherEvent} = $null
        }
        if (${global:Watcher}) {
            ${global:Watcher}.EnableRaisingEvents = $false
            ${global:Watcher}.Dispose()
            ${global:Watcher} = $null
            Write-Log -Message "FileSystemWatcher stopped and disposed." -Level "WARN"
        }
    } catch {
        Write-ErrorLog -Message "Error during FileSystemWatcher cleanup: $($_.Exception.Message)" -Category "FileSystemWatcher" -ErrorRecord $_
    }
}

# Health monitoring and automatic restart
function Monitor-FileSystemWatcher {
    while ($global:ServiceRunning) {
        Start-Sleep -Seconds $global:WatcherHealthCheckInterval
        if (-not (${global:Watcher} -and ${global:Watcher}.EnableRaisingEvents)) {
            Write-Log -Message "FileSystemWatcher not running. Attempting restart..." -Level "WARN"
            try {
                Start-FileSystemWatcher -Path $WatchPath
            } catch {
                Write-ErrorLog -Message "Failed to restart FileSystemWatcher: $($_.Exception.Message)" -Category "FileSystemWatcher" -ErrorRecord $_
            }
        }
    }
}

# Start health monitor in background
Start-Job -ScriptBlock { Monitor-FileSystemWatcher } | Out-Null

# =============================
# Processing Queue (Thread-Safe)
# =============================
# Create a global, thread-safe queue object for audio file processing
$global:Queue = [System.Collections.Queue]::Synchronized((New-Object System.Collections.Queue))
$global:QueueLock = New-Object Object  # Used for thread synchronization

function Test-FileReady {
    param([string]$FilePath)
    $timeout = 300  # seconds
    $elapsed = 0
    $attempt = 0
    $maxAttempts = 10
    while ($elapsed -lt $timeout -and $attempt -lt $maxAttempts) {
        try {
            $stream = [System.IO.File]::Open($FilePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::None)
            $stream.Close()
            return $true
        } catch {
            $backoff = [math]::Min(2 * [math]::Pow(2, $attempt), 30)
            Start-Sleep -Seconds $backoff
            $elapsed += $backoff
            $attempt++
        }
    }
    Write-Log -Message "File not ready after $maxAttempts attempts: $FilePath" -Level "WARN"
    return $false
}

function Add-ToQueue {
    param([string]$FilePath)
    # Use a background job to delay queuing by 15 seconds, then check readiness
    Start-Job -ScriptBlock {
        param($FilePath, $Queue, $QueueLock, $FailedPath)
        Start-Sleep -Seconds 15  # Delay to ensure file is fully written
        if (Test-FileReady -FilePath $FilePath) {
            [System.Threading.Monitor]::Enter($QueueLock)
            try {
                $Queue.Enqueue($FilePath)
                Write-Host "Queued file for processing: $FilePath"
            } finally {
                [System.Threading.Monitor]::Exit($QueueLock)
            }
        } else {
            Write-Host "File not ready after delay and retries: $FilePath. Moving to failed directory."
            try {
                Ensure-Directory $FailedPath
                Move-Item -Path $FilePath -Destination (Join-Path $FailedPath ([System.IO.Path]::GetFileName($FilePath))) -Force
                Write-Log -Message "Moved unprocessable file to failed: $FilePath" -Level "ERROR"
            } catch {
                Write-Log -Message "Failed to move unprocessable file to failed directory: $FilePath. $($_.Exception.Message)" -Level "ERROR"
            }
        }
    } -ArgumentList $FilePath, $global:Queue, $global:QueueLock, $global:FailedPath | Out-Null
}

function Dequeue-File {
    # Thread-safe dequeue operation for the processing queue
    [System.Threading.Monitor]::Enter($global:QueueLock)
    try {
        if ($global:Queue.Count -gt 0) {
            return $global:Queue.Dequeue()
        } else {
            return $null
        }
    } catch {
        Write-Log -Message "Error during dequeue operation: $($_.Exception.Message)" -Level "ERROR"
        return $null
    } finally {
        [System.Threading.Monitor]::Exit($global:QueueLock)
    }
}

# =============================
# Queue Processor Background Job
# =============================
$global:QueueProcessorJob = $null

# =============================
# Whisper CLI Transcription
# =============================
function Ensure-Directory {
    param([string]$Path)
    if (-not (Test-Path $Path)) {
        try {
            New-Item -ItemType Directory -Path $Path -Force | Out-Null
            Write-Log -Message "Created missing directory: $Path" -Level "INFO"
        } catch {
            Write-ErrorLog -Message "Failed to create directory: $Path. $($_.Exception.Message)" -Category "FileManagement" -ErrorRecord $_
            throw
        }
    }
}

function Get-UniqueFileName {
    param([string]$Directory, [string]$BaseName, [string]$Extension)
    $i = 1
    $fileName = "$BaseName$Extension"
    while (Test-Path (Join-Path $Directory $fileName)) {
        $fileName = "$BaseName-$i$Extension"
        $i++
    }
    return (Join-Path $Directory $fileName)
}

function Move-ProcessedFile {
    param(
        [Parameter(Mandatory=$true)][string]$SourceFile,
        [Parameter(Mandatory=$true)][string]$CompletedPath
    )
    try {
        if (-not (Test-Path $SourceFile)) {
            Write-ErrorLog -Message "Source file does not exist: $SourceFile" -Category "FileManagement"
            return $false
        }
        Ensure-Directory $CompletedPath
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($SourceFile)
        $fileExt = [System.IO.Path]::GetExtension($SourceFile)
        $destFile = Get-UniqueFileName -Directory $CompletedPath -BaseName $baseName -Extension $fileExt
        Move-Item -Path $SourceFile -Destination $destFile -Force
        Write-Log -Message "Moved processed file to completed: $destFile" -Level "INFO"
        return $true
    } catch {
        Write-ErrorLog -Message "Failed to move processed file: $($_.Exception.Message)" -Category "FileManagement" -ErrorRecord $_
        return $false
    }
}

function Save-TranscriptFile {
    param(
        [Parameter(Mandatory=$true)][string]$TranscriptContent,
        [Parameter(Mandatory=$true)][string]$BaseName,
        [Parameter(Mandatory=$true)][string]$OutputPath
    )
    try {
        if ([string]::IsNullOrWhiteSpace($TranscriptContent)) {
            Write-ErrorLog -Message "Transcript content is empty or invalid for $BaseName" -Category "Transcript"
            return $null
        }
        Ensure-Directory $OutputPath
        $transcriptFile = Get-UniqueFileName -Directory $OutputPath -BaseName $BaseName -Extension ".txt"
        Set-Content -Path $transcriptFile -Value $TranscriptContent -Encoding UTF8
        Write-Log -Message "Saved transcript file: $transcriptFile" -Level "INFO"
        return $transcriptFile
    } catch {
        Write-ErrorLog -Message "Failed to save transcript file for $BaseName: $($_.Exception.Message)" -Category "Transcript" -ErrorRecord $_
        return $null
    }
}

function Process-AudioFile {
    param(
        [Parameter(Mandatory=$true)][string]$FilePath,
        [string]$Language = 'auto',
        [string]$OutputPath = (Get-Location).Path,
        [string]$Model = 'medium'
    )
    # Validate input file
    if (-not (Test-Path $FilePath)) {
        Write-Log -Message "Audio file does not exist: $FilePath" -Level "ERROR"
        return
    }
    # Validate supported audio formats
    $supportedExtensions = @('.m4a', '.wav', '.mp3')
    $fileExt = [System.IO.Path]::GetExtension($FilePath)
    if ($null -ne $fileExt) { $fileExt = $fileExt.ToLower() } else { $fileExt = '' }
    if ($supportedExtensions -notcontains $fileExt) {
        Write-Log -Message "Unsupported audio format: $fileExt. Supported: $($supportedExtensions -join ', ')" -Level "ERROR"
        return
    }
    # Validate output directory accessibility
    try {
        Ensure-Directory $OutputPath
    } catch {
        Write-Log -Message "Output directory inaccessible: $OutputPath" -Level "ERROR"
        return
    }
    # Input sanitization for file paths (basic)
    if ($FilePath -match '[\r\n\0]') {
        Write-Log -Message "Invalid characters in file path: $FilePath" -Level "ERROR"
        return
    }
    if ($OutputPath -match '[\r\n\0]') {
        Write-Log -Message "Invalid characters in output path: $OutputPath" -Level "ERROR"
        return
    }
    # Validate language code (basic ISO 639-1/2/3 or 'auto')
    if ($Language -ne 'auto' -and $Language -notmatch '^[a-zA-Z]{2,3}$') {
        Write-Log -Message "Invalid language code: $Language. Use ISO 639-1/2/3 or 'auto'." -Level "ERROR"
        return
    }
    Ensure-Directory $CompletedPath
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
    $outputFile = Get-UniqueFileName -Directory $OutputPath -BaseName $baseName -Extension ".txt"
    $whisperArgs = @(
        "--model", $Model,
        "--language", $Language,
        "--output_format", "txt",
        "--output_dir", $OutputPath,
        $FilePath
    )
    # Check if whisper CLI is available
    $whisperCmd = Get-Command whisper -ErrorAction SilentlyContinue
    if (-not $whisperCmd) {
        Write-ErrorLog -Message "Whisper CLI executable not found in PATH. Please install or add to PATH." -Category "WhisperCLI"
        return
    }
    $maxRetries = 2
    $retryCount = 0
    $success = $false
    $pauseOnCriticalError = $false
    while (-not $success -and $retryCount -le $maxRetries) {
        try {
            $psi = New-Object System.Diagnostics.ProcessStartInfo
            $psi.FileName = $whisperCmd.Source
            $psi.Arguments = $whisperArgs -join ' '
            $psi.RedirectStandardOutput = $true
            $psi.RedirectStandardError = $true
            $psi.UseShellExecute = $false
            $psi.CreateNoWindow = $true
            $process = [System.Diagnostics.Process]::Start($psi)
            # Timeout handling (10 minutes = 600 seconds)
            $timeout = 600
            $waited = 0
            while (-not $process.HasExited -and $waited -lt $timeout) {
                Start-Sleep -Seconds 2
                $waited += 2
            }
            if (-not $process.HasExited) {
                $process.Kill()
                Write-ErrorLog -Message "Whisper CLI process timed out after $timeout seconds for $FilePath." -Category "WhisperCLI"
                $retryCount++
                continue
            }
            $stdout = $process.StandardOutput.ReadToEnd()
            $stderr = $process.StandardError.ReadToEnd()
            if ($process.ExitCode -eq 0) {
                # Save transcript file if stdout is not empty
                $savedTranscript = $null
                if (-not [string]::IsNullOrWhiteSpace($stdout)) {
                    $savedTranscript = Save-TranscriptFile -TranscriptContent $stdout -BaseName $baseName -OutputPath $OutputPath
                }
                # Move original audio file to CompletedPath using new function
                $moved = Move-ProcessedFile -SourceFile $FilePath -CompletedPath $CompletedPath
                if (-not $moved) {
                    Write-ErrorLog -Message "Failed to move processed file after transcription: $FilePath" -Category "FileManagement"
                }
                Write-Log -Message "Transcription complete: $outputFile" -Level "INFO"
                $success = $true
            } else {
                # Handle specific error scenarios
                if ($stderr -match 'No space left on device') {
                    Write-ErrorLog -Message "Insufficient disk space for $FilePath: $stderr" -Category "WhisperCLI"
                    $pauseOnCriticalError = $true
                    break
                } elseif ($stderr -match 'Permission denied') {
                    Write-ErrorLog -Message "Permission denied for $FilePath: $stderr" -Category "WhisperCLI"
                    break
                } elseif ($stderr -match 'invalid' -or $stderr -match 'unsupported') {
                    Write-ErrorLog -Message "Invalid or unsupported audio file: $FilePath. $stderr" -Category "WhisperCLI"
                    break
                } elseif ($stderr -match 'network|connection|timeout|unreachable') {
                    Write-ErrorLog -Message "Network error for $FilePath: $stderr" -Category "Network"
                    $backoff = [math]::Min(5 * [math]::Pow(2, $retryCount), 60)
                    Start-Sleep -Seconds $backoff
                    $retryCount++
                    continue
                } elseif ($stderr -match 'memory|out of memory|insufficient memory') {
                    Write-ErrorLog -Message "Memory exhaustion for $FilePath: $stderr" -Category "Memory"
                    $pauseOnCriticalError = $true
                    break
                } elseif ($stderr -match 'error' -or $stderr -match 'fail' -or $stderr -match 'exception') {
                    Write-ErrorLog -Message "Whisper CLI error for $FilePath: $stderr" -Category "WhisperCLI"
                    $backoff = [math]::Min(5 * [math]::Pow(2, $retryCount), 60)
                    Start-Sleep -Seconds $backoff
                    $retryCount++
                    continue
                } else {
                    Write-ErrorLog -Message "Whisper CLI failed for $FilePath: $stderr" -Category "WhisperCLI"
                    $backoff = [math]::Min(5 * [math]::Pow(2, $retryCount), 60)
                    Start-Sleep -Seconds $backoff
                    $retryCount++
                    continue
                }
            }
        } catch {
            Write-ErrorLog -Message "Exception in Process-AudioFile for $FilePath: $($_.Exception.Message)" -Category "WhisperCLI" -ErrorRecord $_
            $backoff = [math]::Min(5 * [math]::Pow(2, $retryCount), 60)
            Start-Sleep -Seconds $backoff
            $retryCount++
            continue
        } finally {
            # Cleanup: remove any temp files if needed (add logic if temp files are used)
        }
    }
    if (-not $success) {
        Write-ErrorLog -Message "Whisper CLI failed after $maxRetries retries for $FilePath. Moving to failed directory." -Category "WhisperCLI"
        try {
            Ensure-Directory $global:FailedPath
            Move-Item -Path $FilePath -Destination (Join-Path $global:FailedPath ([System.IO.Path]::GetFileName($FilePath))) -Force
            Write-Log -Message "Moved failed file to failed directory: $FilePath" -Level "ERROR"
        } catch {
            Write-Log -Message "Failed to move failed file to failed directory: $FilePath. $($_.Exception.Message)" -Level "ERROR"
        }
    }
    if ($pauseOnCriticalError) {
        Write-Log -Message "Pausing queue processor for 5 minutes due to critical error (disk/memory)." -Level "ERROR"
        Start-Sleep -Seconds 300
    }
}

function Start-QueueProcessor {
    if ($global:QueueProcessorJob -and ($global:QueueProcessorJob.State -eq 'Running')) {
        Write-Log -Message "Queue processor already running." -Level "WARN"
        return
    }
    $global:QueueProcessorJob = Start-Job -ScriptBlock {
        param($LogPath, $CheckInterval, $OutputPath, $Language, $Queue, $QueueLock)
        Import-Module Microsoft.PowerShell.Utility
        function Write-Log { param($Message, $Level) Write-Host "[$Level] $Message" }
        function Write-ErrorLog { param($Message, $Category) Write-Host "[ERROR][$Category] $Message" }
        function Dequeue-File {
            param($Queue, $QueueLock)
            [System.Threading.Monitor]::Enter($QueueLock)
            try {
                if ($Queue.Count -gt 0) {
                    return $Queue.Dequeue()
                } else {
                    return $null
                }
            } catch {
                Write-Log "Error during dequeue operation: $($_.Exception.Message)" "ERROR"
                return $null
            } finally {
                [System.Threading.Monitor]::Exit($QueueLock)
            }
        }
        function Process-AudioFile {
            param([string]$FilePath)
            if (-not (Test-Path $FilePath)) {
                Write-Log "Audio file does not exist: $FilePath" "ERROR"
                return
            }
            $baseName = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
            $outputFile = Join-Path $OutputPath ("$baseName.txt")
            $whisperArgs = @(
                "--model", "medium",
                "--language", $Language,
                "--output_format", "txt",
                "--output_dir", $OutputPath,
                $FilePath
            )
            try {
                $psi = New-Object System.Diagnostics.ProcessStartInfo
                $psi.FileName = "whisper"
                $psi.Arguments = $whisperArgs -join ' '
                $psi.RedirectStandardOutput = $true
                $psi.RedirectStandardError = $true
                $psi.UseShellExecute = $false
                $psi.CreateNoWindow = $true
                $process = [System.Diagnostics.Process]::Start($psi)
                $stdout = $process.StandardOutput.ReadToEnd()
                $stderr = $process.StandardError.ReadToEnd()
                $process.WaitForExit()
                if ($process.ExitCode -eq 0) {
                    Write-Log "Transcription complete: $outputFile" "INFO"
                } else {
                    Write-ErrorLog "Whisper CLI failed for $FilePath: $stderr" "WhisperCLI"
                }
            } catch {
                Write-ErrorLog "Exception in Process-AudioFile for $FilePath: $($_.Exception.Message)" "WhisperCLI"
            }
        }
        while ($true) {
            try {
                $file = Dequeue-File $Queue $QueueLock
                if ($file) {
                    Write-Log "Processing file from queue: $file" "INFO"
                    Process-AudioFile $file
                } else {
                    Start-Sleep -Seconds $CheckInterval
                }
            } catch {
                Write-Log "Error in queue processor: $($_.Exception.Message)" "ERROR"
            }
        }
    } -ArgumentList $LogPath, $CheckInterval, $OutputPath, $Language, $global:Queue, $global:QueueLock
    Write-Log -Message "Queue processor job started." -Level "INFO"
}

function Stop-QueueProcessor {
    if ($global:QueueProcessorJob) {
        Stop-Job $global:QueueProcessorJob -Force -ErrorAction SilentlyContinue
        Remove-Job $global:QueueProcessorJob -Force -ErrorAction SilentlyContinue
        $global:QueueProcessorJob = $null
        Write-Log -Message "Queue processor job stopped." -Level "WARN"
    }
}

function Restart-QueueProcessor {
    Stop-QueueProcessor
    Start-QueueProcessor
}

# Start the queue processor with the service
Start-QueueProcessor

# Add queue processor cleanup to service shutdown
$null = Register-EngineEvent PowerShell.Exiting -Action {
    $global:ServiceRunning = $false
    Stop-QueueProcessor
    Stop-FileSystemWatcher
    Write-Log -Message "Service shutdown signal received. Cleaning up..." -Level "WARN"
} 

function Initialize-StartupFileProcessing {
    param(
        [string]$WatchPath,
        [string]$QueueLock,
        $Queue
    )
    Write-Log -Message "Scanning for existing .m4a files in $WatchPath on startup..." -Level "INFO"
    try {
        $files = Get-ChildItem -Path $WatchPath -Filter *.m4a -File -ErrorAction SilentlyContinue
        if ($files.Count -eq 0) {
            Write-Log -Message "No existing .m4a files found in $WatchPath." -Level "INFO"
            return
        }
        foreach ($file in $files) {
            $filePath = $file.FullName
            if (Test-FileReady -FilePath $filePath) {
                Add-ToQueue -FilePath $filePath
                Write-Log -Message "Queued existing file for processing: $filePath" -Level "INFO"
            } else {
                Write-Log -Message "File not ready (locked or inaccessible): $filePath" -Level "WARN"
            }
        }
        Write-Log -Message "Startup file scan complete. Queue count: $($Queue.Count)" -Level "INFO"
    } catch {
        Write-Log -Message "Error during startup file scan: $($_.Exception.Message)" -Level "ERROR"
    }
}

# Call Initialize-StartupFileProcessing at service startup
Initialize-StartupFileProcessing -WatchPath $WatchPath -QueueLock $global:QueueLock -Queue $global:Queue 

function Restart-FailedComponents {
    param(
        [int]$MaxAttempts = 5
    )
    $attempt = 0
    $backoff = 5
    while ($attempt -lt $MaxAttempts) {
        $restarted = $false
        # Restart FileSystemWatcher if needed
        if ($null -eq $global:Watcher -or $global:Watcher.EnableRaisingEvents -ne $true) {
            try {
                Write-Log -Message "Restarting FileSystemWatcher... (attempt $($attempt+1))" -Level "WARN"
                Start-FileSystemWatcher -Path $WatchPath
                $restarted = $true
            } catch {
                Write-Log -Message "Failed to restart FileSystemWatcher: $($_.Exception.Message)" -Level "ERROR"
            }
        }
        # Restart QueueProcessorJob if needed
        if ($null -eq $global:QueueProcessorJob -or $global:QueueProcessorJob.State -ne 'Running') {
            try {
                Write-Log -Message "Restarting QueueProcessorJob... (attempt $($attempt+1))" -Level "WARN"
                Restart-QueueProcessor
                $restarted = $true
            } catch {
                Write-Log -Message "Failed to restart QueueProcessorJob: $($_.Exception.Message)" -Level "ERROR"
            }
        }
        if ($restarted) {
            Write-Log -Message "Component(s) restarted successfully." -Level "INFO"
            break
        } else {
            $attempt++
            Start-Sleep -Seconds ($backoff * $attempt)
        }
    }
    if ($attempt -ge $MaxAttempts) {
        Write-Log -Message "Max restart attempts reached. Manual intervention required." -Level "ERROR"
    }
}

# Update health monitor to call Restart-FailedComponents if needed
function Start-HealthMonitor {
    param(
        [int]$IntervalSeconds = 60
    )
    if ($global:HealthMonitorJob -and ($global:HealthMonitorJob.State -eq 'Running')) {
        Write-Log -Message "Health monitor already running." -Level "WARN"
        return
    }
    $global:HealthMonitorJob = Start-Job -ScriptBlock {
        param($IntervalSeconds)
        while ($true) {
            $needsRestart = $false
            try {
                # Check FileSystemWatcher
                if ($null -eq $global:Watcher -or $global:Watcher.EnableRaisingEvents -ne $true) {
                    Write-Host "[ERROR][Health] FileSystemWatcher is not active!" -ForegroundColor Red
                    $needsRestart = $true
                }
                # Check Queue Processor Job
                if ($null -eq $global:QueueProcessorJob -or $global:QueueProcessorJob.State -ne 'Running') {
                    Write-Host "[ERROR][Health] Queue processor job is not running!" -ForegroundColor Red
                    $needsRestart = $true
                }
                # System resource metrics
                $mem = [math]::Round((Get-Process -Id $PID).WorkingSet64 / 1MB, 2)
                $queueDepth = if ($null -ne $global:Queue) { $global:Queue.Count } else { 0 }
                Write-Host "[HEALTH] Memory: ${mem}MB, Queue Depth: $queueDepth" -ForegroundColor Cyan
            } catch {
                Write-Host "[ERROR][Health] Exception in health monitor: $($_.Exception.Message)" -ForegroundColor Red
                $needsRestart = $true
            }
            if ($needsRestart) {
                Restart-FailedComponents
            }
            Start-Sleep -Seconds $IntervalSeconds
        }
    } -ArgumentList $IntervalSeconds
    Write-Log -Message "Health monitor job started." -Level "INFO"
}

# Start health monitor at service startup
Start-HealthMonitor -IntervalSeconds 60 

function Stop-ServiceGracefully {
    Write-Log -Message "Initiating graceful shutdown..." -Level "WARN"
    try {
        # Persist queue state to disk for recovery
        try {
            $queueStatePath = Join-Path $LogPath 'queue-state.json'
            $queueData = if ($null -ne $global:Queue) { $global:Queue | ConvertTo-Json -Depth 5 } else { '[]' }
            Set-Content -Path $queueStatePath -Value $queueData -Encoding UTF8
            Write-Log -Message "Queue state persisted to $queueStatePath." -Level "INFO"
        } catch {
            Write-Log -Message "Error persisting queue state: $($_.Exception.Message)" -Level "ERROR"
        }

        # Stop FileSystemWatcher with timeout
        if ($global:Watcher) {
            $watcherStopped = $false
            $watcherJob = Start-Job -ScriptBlock {
                param($Watcher)
                try {
                    $Watcher.EnableRaisingEvents = $false
                    $Watcher.Dispose()
                    return $true
                } catch { return $false }
            } -ArgumentList $global:Watcher
            if (Wait-Job $watcherJob -Timeout 10) {
                $watcherStopped = Receive-Job $watcherJob
                Write-Log -Message "FileSystemWatcher stopped and disposed." -Level "INFO"
            } else {
                Write-Log -Message "Timeout stopping FileSystemWatcher. Forcing termination." -Level "ERROR"
                Stop-Job $watcherJob -Force
            }
            Remove-Job $watcherJob -Force
        }

        # Stop Queue Processor Job with timeout
        if ($global:QueueProcessorJob) {
            try {
                $jobId = $global:QueueProcessorJob.Id
                Stop-Job $global:QueueProcessorJob -Force -ErrorAction SilentlyContinue
                if (-not (Wait-Job $global:QueueProcessorJob -Timeout 10)) {
                    Write-Log -Message "Timeout stopping queue processor job. Forcing removal." -Level "ERROR"
                }
                Remove-Job $global:QueueProcessorJob -Force -ErrorAction SilentlyContinue
                Write-Log -Message "Queue processor job stopped and removed." -Level "INFO"
            } catch {
                Write-Log -Message "Error stopping queue processor job: $($_.Exception.Message)" -Level "ERROR"
            }
        }

        # Stop Health Monitor Job with timeout
        if ($global:HealthMonitorJob) {
            try {
                Stop-Job $global:HealthMonitorJob -Force -ErrorAction SilentlyContinue
                if (-not (Wait-Job $global:HealthMonitorJob -Timeout 10)) {
                    Write-Log -Message "Timeout stopping health monitor job. Forcing removal." -Level "ERROR"
                }
                Remove-Job $global:HealthMonitorJob -Force -ErrorAction SilentlyContinue
                Write-Log -Message "Health monitor job stopped and removed." -Level "INFO"
            } catch {
                Write-Log -Message "Error stopping health monitor job: $($_.Exception.Message)" -Level "ERROR"
            }
        }

        # Flush log buffers (if using buffered logging, add logic here)
        try {
            [System.GC]::Collect() # Encourage cleanup of any open handles
            Write-Log -Message "Log buffers flushed (if applicable)." -Level "INFO"
        } catch {
            Write-Log -Message "Error flushing log buffers: $($_.Exception.Message)" -Level "ERROR"
        }

        # Additional resource cleanup placeholder
        # e.g., Remove temporary files, close custom handles, etc.
        # Add any additional cleanup logic here

        Write-Log -Message "Graceful shutdown complete." -Level "INFO"
    } catch {
        Write-Log -Message "Exception during graceful shutdown: $($_.Exception.Message)" -Level "ERROR"
    }
}

# Register graceful shutdown handler
$null = Register-EngineEvent PowerShell.Exiting -Action { Stop-ServiceGracefully } 

function Start-ServiceController {
    Write-Log -Message "[ServiceController] Starting service lifecycle management..." -Level "INFO"
    try {
        # 1. Configuration Validation
        Write-Log -Message "[ServiceController] Validating configuration..." -Level "INFO"
        if (-not (Test-Path $WatchPath)) {
            Write-Log -Message "[ServiceController] WatchPath does not exist: $WatchPath" -Level "ERROR"
            throw "WatchPath does not exist: $WatchPath"
        }
        Ensure-Directory $OutputPath
        Ensure-Directory $CompletedPath
        Ensure-Directory $LogPath
        Ensure-Directory $global:FailedPath

        # 2. Initialize Logging
        Write-Log -Message "[ServiceController] Logging initialized." -Level "INFO"

        # 3. Initialize Queue and Lock
        if (-not $global:Queue) {
            $global:Queue = [System.Collections.Queue]::Synchronized((New-Object System.Collections.Queue))
            Write-Log -Message "[ServiceController] Queue initialized." -Level "INFO"
        }
        if (-not $global:QueueLock) {
            $global:QueueLock = New-Object Object
            Write-Log -Message "[ServiceController] QueueLock initialized." -Level "INFO"
        }

        # 4. Startup File Scan
        Write-Log -Message "[ServiceController] Scanning for existing files at startup..." -Level "INFO"
        Initialize-StartupFileProcessing -WatchPath $WatchPath -QueueLock $global:QueueLock -Queue $global:Queue

        # 5. Start FileSystemWatcher
        Write-Log -Message "[ServiceController] Starting FileSystemWatcher..." -Level "INFO"
        Start-FileSystemWatcher -Path $WatchPath

        # 6. Start Queue Processor
        Write-Log -Message "[ServiceController] Starting Queue Processor..." -Level "INFO"
        Start-QueueProcessor

        # 7. Start Health Monitor
        Write-Log -Message "[ServiceController] Starting Health Monitor..." -Level "INFO"
        Start-HealthMonitor -IntervalSeconds 60

        # 8. Register Graceful Shutdown
        Write-Log -Message "[ServiceController] Registering graceful shutdown handler..." -Level "INFO"
        $null = Register-EngineEvent PowerShell.Exiting -Action { Stop-ServiceGracefully }

        Write-Log -Message "[ServiceController] Service lifecycle management started successfully." -Level "INFO"
    } catch {
        Write-Log -Message "[ServiceController] Exception during service startup: $($_.Exception.Message)" -Level "ERROR"
        throw
    }
}

# --- Call the master service controller at the end of the script ---
Start-ServiceController 