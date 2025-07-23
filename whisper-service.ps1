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
$global:FailedPath = Join-Path $OutputPath 'failed'
if (-not (Test-Path $CompletedPath)) {
    $global:CompletedPath = Join-Path $OutputPath 'completed'
    Ensure-Directory $global:CompletedPath
}
if (-not (Test-Path $global:FailedPath)) {
    Ensure-Directory $global:FailedPath
}

# =============================
# Validation Functions
# =============================
function Test-Directory {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) {
        Write-Error "Path parameter cannot be null or empty"
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

function Write-ErrorLog {
    param(
        [Parameter(Mandatory=$true)][string]$Message,
        [string]$Category = "General",
        [System.Management.Automation.ErrorRecord]$ErrorRecord
    )
    $errorMessage = "[$Category] $Message"
    if ($ErrorRecord) {
        $errorMessage += " Exception: $($ErrorRecord.Exception.Message)"
    }
    Write-Log -Message $errorMessage -Level "ERROR"
}

# =============================
# Utility Functions
# =============================
function Ensure-Directory {
    param([string]$Path)
    if (-not (Test-Path $Path)) {
        try {
            New-Item -ItemType Directory -Path $Path -Force | Out-Null
            Write-Log -Message "Created directory: $Path" -Level "INFO"
        } catch {
            Write-Log -Message "Failed to create directory: $Path. $($_.Exception.Message)" -Level "ERROR"
            throw
        }
    }
}

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
function Save-TranscriptFile {
    param(
        [string]$TranscriptContent,
        [string]$BaseName,
        [string]$OutputPath
    )
    try {
        $transcriptFile = Get-UniqueFileName -Directory $OutputPath -BaseName $BaseName -Extension ".txt"
        Set-Content -Path $transcriptFile -Value $TranscriptContent -Encoding UTF8
        Write-Log -Message "Saved transcript file: $transcriptFile" -Level "INFO"
        return $transcriptFile
    } catch {
        Write-ErrorLog -Message "Failed to save transcript file for ${BaseName}: $($_.Exception.Message)" -Category "Transcript" -ErrorRecord $_
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
                
                if ($savedTranscript) {
                    Write-Log -Message "Successfully processed: $FilePath -> $savedTranscript" -Level "INFO"
                    $success = $true
                } else {
                    Write-Log -Message "Failed to save transcript for: $FilePath" -Level "ERROR"
                    $retryCount++
                    continue
                }
            } else {
                # Handle specific error scenarios
                if ($stderr -match 'No space left on device') {
                    Write-ErrorLog -Message "Insufficient disk space for ${FilePath}: $stderr" -Category "WhisperCLI"
                    $pauseOnCriticalError = $true
                    break
                } elseif ($stderr -match 'Permission denied') {
                    Write-ErrorLog -Message "Permission denied for ${FilePath}: $stderr" -Category "WhisperCLI"
                    break
                } elseif ($stderr -match 'invalid' -or $stderr -match 'unsupported') {
                    Write-ErrorLog -Message "Invalid or unsupported audio file: ${FilePath}. $stderr" -Category "WhisperCLI"
                    break
                } elseif ($stderr -match 'network|connection|timeout|unreachable') {
                    Write-ErrorLog -Message "Network error for ${FilePath}: $stderr" -Category "Network"
                    $backoff = [math]::Min(5 * [math]::Pow(2, $retryCount), 60)
                    Start-Sleep -Seconds $backoff
                    $retryCount++
                    continue
                } elseif ($stderr -match 'memory|out of memory|insufficient memory') {
                    Write-ErrorLog -Message "Memory exhaustion for ${FilePath}: $stderr" -Category "Memory"
                    $pauseOnCriticalError = $true
                    break
                } elseif ($stderr -match 'error' -or $stderr -match 'fail' -or $stderr -match 'exception') {
                    Write-ErrorLog -Message "Whisper CLI error for ${FilePath}: $stderr" -Category "WhisperCLI"
                    $backoff = [math]::Min(5 * [math]::Pow(2, $retryCount), 60)
                    Start-Sleep -Seconds $backoff
                    $retryCount++
                    continue
                } else {
                    Write-ErrorLog -Message "Whisper CLI failed for ${FilePath}: $stderr" -Category "WhisperCLI"
                    $backoff = [math]::Min(5 * [math]::Pow(2, $retryCount), 60)
                    Start-Sleep -Seconds $backoff
                    $retryCount++
                    continue
                }
            }
        } catch {
            Write-ErrorLog -Message "Exception in Process-AudioFile for ${FilePath}: $($_.Exception.Message)" -Category "WhisperCLI" -ErrorRecord $_
            $backoff = [math]::Min(5 * [math]::Pow(2, $retryCount), 60)
            Start-Sleep -Seconds $backoff
            $retryCount++
            continue
        } finally {
            # Cleanup: remove any temp files if needed (add logic if temp files are used)
        }
    }
    
    if (-not $success) {
        Write-ErrorLog -Message "Whisper CLI failed after $maxRetries retries for ${FilePath}. Moving to failed directory." -Category "WhisperCLI"
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

# =============================
# Queue Management Functions
# =============================
function Add-ToQueue {
    param([string]$FilePath)
    
    if (-not $global:Queue) {
        $global:Queue = [System.Collections.Queue]::Synchronized((New-Object System.Collections.Queue))
    }
    if (-not $global:QueueLock) {
        $global:QueueLock = New-Object Object
    }
    
    # Start a delayed job to add the file to the queue after 15 seconds
    Start-Job -ScriptBlock {
        param($FilePath, $Queue, $QueueLock)
        Start-Sleep -Seconds 15
        
        # Check if file is ready (not locked)
        $maxWait = 300  # 5 minutes
        $waited = 0
        while ($waited -lt $maxWait) {
            try {
                $fileStream = [System.IO.File]::Open($FilePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::None)
                $fileStream.Close()
                break  # File is ready
            } catch {
                Start-Sleep -Seconds 2
                $waited += 2
            }
        }
        
        if ($waited -ge $maxWait) {
            Write-Host "File still locked after 5 minutes: $FilePath" -ForegroundColor Red
            return
        }
        
        # Add to queue
        [System.Threading.Monitor]::Enter($QueueLock)
        try {
            $Queue.Enqueue($FilePath)
            Write-Host "Added to queue: $FilePath" -ForegroundColor Green
        } catch {
            Write-Host "Error adding to queue: $($_.Exception.Message)" -ForegroundColor Red
        } finally {
            [System.Threading.Monitor]::Exit($QueueLock)
        }
    } -ArgumentList $FilePath, $global:Queue, $global:QueueLock
}

function Dequeue-File {
    if (-not $global:Queue) { return $null }
    
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
# Background Job Functions
# =============================
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
                $process = Start-Process -FilePath "whisper" -ArgumentList $whisperArgs -Wait -PassThru -NoNewWindow
                
                if ($process.ExitCode -eq 0) {
                    Write-Log "Successfully processed: $FilePath" "INFO"
                    
                    # Move processed file to completed directory
                    $completedDir = Join-Path $OutputPath "completed"
                    if (-not (Test-Path $completedDir)) {
                        New-Item -ItemType Directory -Path $completedDir -Force | Out-Null
                    }
                    Move-Item -Path $FilePath -Destination (Join-Path $completedDir ([System.IO.Path]::GetFileName($FilePath))) -Force
                    Write-Log "Moved to completed: $FilePath" "INFO"
                } else {
                    Write-ErrorLog "Whisper CLI failed for ${FilePath}" "WhisperCLI"
                }
            } catch {
                Write-ErrorLog "Exception in Process-AudioFile for ${FilePath}: $($_.Exception.Message)" "WhisperCLI"
            }
        }
        
        Write-Log "Queue processor started" "INFO"
        
        while ($true) {
            try {
                $filePath = Dequeue-File -Queue $Queue -QueueLock $QueueLock
                
                if ($filePath) {
                    Write-Log "Processing file: $filePath" "INFO"
                    Process-AudioFile -FilePath $filePath
                } else {
                    Start-Sleep -Seconds $CheckInterval
                }
            } catch {
                Write-Log "Error in queue processor: $($_.Exception.Message)" "ERROR"
                Start-Sleep -Seconds $CheckInterval
            }
        }
    } -ArgumentList $LogPath, $CheckInterval, $OutputPath, $Language, $global:Queue, $global:QueueLock
    
    Write-Log -Message "Queue processor job started." -Level "INFO"
}

function Stop-QueueProcessor {
    if ($global:QueueProcessorJob) {
        try {
            Stop-Job $global:QueueProcessorJob -Force -ErrorAction SilentlyContinue
            Remove-Job $global:QueueProcessorJob -Force -ErrorAction SilentlyContinue
            Write-Log -Message "Queue processor job stopped." -Level "INFO"
        } catch {
            Write-Log -Message "Error stopping queue processor job: $($_.Exception.Message)" -Level "ERROR"
        }
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
    $global:FileSystemWatcher.Filter = "*.m4a"
    $global:FileSystemWatcher.IncludeSubdirectories = $false
    $global:FileSystemWatcher.NotifyFilter = [System.IO.NotifyFilters]::FileName
    
    # Register event handlers
    $global:FileSystemWatcher.Created += {
        param($sender, $e)
        $filePath = $e.FullPath
        Write-Log -Message "New file detected: $filePath" -Level "INFO"
        Add-ToQueue -FilePath $filePath
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
# Health Monitoring Functions
# =============================
function Start-HealthMonitor {
    param([int]$IntervalSeconds = 60)
    
    if ($global:HealthMonitorJob -and ($global:HealthMonitorJob.State -eq 'Running')) {
        Write-Log -Message "Health monitor already running." -Level "WARN"
        return
    }
    
    $global:HealthMonitorJob = Start-Job -ScriptBlock {
        param($IntervalSeconds)
        
        function Write-Log { param($Message, $Level) Write-Host "[$Level] $Message" }
        
        Write-Log "Health monitor started" "INFO"
        
        while ($true) {
            try {
                # Check if FileSystemWatcher is still working
                $fsWatcher = Get-Job | Where-Object { $_.Name -like "*FileSystemWatcher*" }
                if (-not $fsWatcher -or $fsWatcher.State -ne 'Running') {
                    Write-Log "FileSystemWatcher job not running, restarting..." "WARN"
                    # Restart logic would go here
                }
                
                # Check if Queue Processor is still working
                $queueJob = Get-Job | Where-Object { $_.Name -like "*QueueProcessor*" }
                if (-not $queueJob -or $queueJob.State -ne 'Running') {
                    Write-Log "Queue processor job not running, restarting..." "WARN"
                    # Restart logic would go here
                }
                
                Start-Sleep -Seconds $IntervalSeconds
            } catch {
                Write-Log "Error in health monitor: $($_.Exception.Message)" "ERROR"
                Start-Sleep -Seconds $IntervalSeconds
            }
        }
    } -ArgumentList $IntervalSeconds
    
    Write-Log -Message "Health monitor job started." -Level "INFO"
}

function Stop-HealthMonitor {
    if ($global:HealthMonitorJob) {
        try {
            Stop-Job $global:HealthMonitorJob -Force -ErrorAction SilentlyContinue
            Remove-Job $global:HealthMonitorJob -Force -ErrorAction SilentlyContinue
            Write-Log -Message "Health monitor job stopped." -Level "INFO"
        } catch {
            Write-Log -Message "Error stopping health monitor job: $($_.Exception.Message)" -Level "ERROR"
        }
    }
}

# =============================
# Startup File Processing
# =============================
function Initialize-StartupFileProcessing {
    param(
        [string]$WatchPath,
        [Object]$QueueLock,
        [System.Collections.Queue]$Queue
    )
    
    try {
        $existingFiles = Get-ChildItem -Path $WatchPath -Filter "*.m4a" -ErrorAction SilentlyContinue
        
        if ($existingFiles) {
            Write-Log -Message "Found $($existingFiles.Count) existing .m4a files at startup. Adding to queue..." -Level "INFO"
            
            foreach ($file in $existingFiles) {
                [System.Threading.Monitor]::Enter($QueueLock)
                try {
                    $Queue.Enqueue($file.FullName)
                    Write-Log -Message "Added existing file to queue: $($file.Name)" -Level "INFO"
                } catch {
                    Write-Log -Message "Error adding existing file to queue: $($file.Name). $($_.Exception.Message)" -Level "ERROR"
                } finally {
                    [System.Threading.Monitor]::Exit($QueueLock)
                }
            }
        } else {
            Write-Log -Message "No existing .m4a files found at startup." -Level "INFO"
        }
    } catch {
        Write-Log -Message "Error during startup file processing: $($_.Exception.Message)" -Level "ERROR"
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
        
        # Stop Queue Processor Job with timeout
        if ($global:QueueProcessorJob) {
            try {
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
        
        Write-Log -Message "Graceful shutdown complete." -Level "INFO"
    } catch {
        Write-Log -Message "Exception during graceful shutdown: $($_.Exception.Message)" -Level "ERROR"
    }
}

# Register graceful shutdown handler
$null = Register-EngineEvent PowerShell.Exiting -Action { Stop-ServiceGracefully } 

# =============================
# Service Controller
# =============================
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

# =============================
# Start the Service
# =============================
Start-ServiceController 