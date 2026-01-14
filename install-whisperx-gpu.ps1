# install-whisperx-gpu.ps1
# WhisperX GPU Installation Script for NVIDIA RTX GPUs
# This script automates the installation of WhisperX with proper GPU support
#
# Requirements:
# - Windows 10/11
# - NVIDIA RTX GPU (RTX 2000 series or newer recommended)
# - NVIDIA drivers installed
# - Python 3.10+ installed
# - PowerShell 5.1 or newer

# =============================
# Configuration Parameters
# =============================
param(
    [string]$InstallPath = "C:\whisper",
    [string]$PythonVersion = "3.12",
    [switch]$SkipNvidiaCheck = $false,
    [switch]$Force = $false
)

# =============================
# Logging Functions
# =============================
function Write-InstallLog {
    param(
        [Parameter(Mandatory=$true)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR','SUCCESS')][string]$Level = 'INFO'
    )
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $colors = @{
        'INFO' = 'White'
        'WARN' = 'Yellow'
        'ERROR' = 'Red'
        'SUCCESS' = 'Green'
    }
    $logLine = "[$timestamp] [$Level] $Message"
    Write-Host $logLine -ForegroundColor $colors[$Level]
}

# =============================
# System Validation Functions
# =============================
function Test-NvidiaGPU {
    Write-InstallLog "Checking for NVIDIA GPU..." -Level "INFO"
    try {
        $nvidiaOutput = nvidia-smi 2>$null
        if ($LASTEXITCODE -eq 0) {
            $gpuInfo = nvidia-smi --query-gpu=name,driver_version,memory.total --format=csv,noheader,nounits 2>$null
            if ($gpuInfo) {
                Write-InstallLog "NVIDIA GPU detected: $($gpuInfo.Split(',')[0].Trim())" -Level "SUCCESS"
                $driverVersion = $gpuInfo.Split(',')[1].Trim()
                Write-InstallLog "Driver version: $driverVersion" -Level "INFO"
                return $true
            }
        }
    } catch {
        Write-InstallLog "nvidia-smi not found or failed" -Level "WARN"
    }
    return $false
}

function Test-PythonVersion {
    Write-InstallLog "Checking Python installation..." -Level "INFO"
    try {
        $pythonVersion = python --version 2>&1
        if ($pythonVersion -match "Python (\d+\.\d+)") {
            $version = [version]$matches[1]
            Write-InstallLog "Python version: $pythonVersion" -Level "INFO"
            if ($version -ge [version]"3.10") {
                Write-InstallLog "Python version is compatible" -Level "SUCCESS"
                return $true
            } else {
                Write-InstallLog "Python version must be 3.10 or newer" -Level "ERROR"
                return $false
            }
        }
    } catch {
        Write-InstallLog "Python not found in PATH" -Level "ERROR"
        return $false
    }
    return $false
}

function Test-PyTorchCUDA {
    Write-InstallLog "Checking PyTorch CUDA support..." -Level "INFO"
    try {
        $cudaCheck = python -c "import torch; print(f'PyTorch: {torch.__version__}'); print(f'CUDA: {torch.cuda.is_available()}')" 2>$null
        if ($cudaCheck) {
            Write-InstallLog "PyTorch CUDA status: $cudaCheck" -Level "INFO"
            if ($cudaCheck -match "CUDA: True") {
                Write-InstallLog "PyTorch CUDA support confirmed" -Level "SUCCESS"
                return $true
            }
        }
    } catch {
        Write-InstallLog "PyTorch not installed or CUDA not available" -Level "WARN"
    }
    return $false
}

# =============================
# Installation Functions
# =============================
function Install-PyTorchCUDA {
    Write-InstallLog "Installing PyTorch with CUDA 11.8 support..." -Level "INFO"
    try {
        $result = pip install torch torchvision torchaudio --index-url https://download.pytorch.org/whl/cu118 --upgrade
        if ($LASTEXITCODE -eq 0) {
            Write-InstallLog "PyTorch CUDA installation completed" -Level "SUCCESS"
            return $true
        } else {
            Write-InstallLog "PyTorch installation failed" -Level "ERROR"
            return $false
        }
    } catch {
        Write-InstallLog "Exception during PyTorch installation: $_" -Level "ERROR"
        return $false
    }
}

function Install-CUDNNLibraries {
    Write-InstallLog "Installing cuDNN libraries..." -Level "INFO"
    try {
        # Install cuDNN 8.x compatible with CUDA 11.8
        pip install nvidia-cudnn-cu11==8.9.4.25 --upgrade
        if ($LASTEXITCODE -eq 0) {
            Write-InstallLog "cuDNN 8.9.4.25 installation completed" -Level "SUCCESS"
        } else {
            Write-InstallLog "cuDNN installation failed" -Level "ERROR"
            return $false
        }

        # Verify cuDNN installation
        $pythonSitePackages = python -c "import site; print(site.getsitepackages()[1])" 2>$null
        $cudnnPath = Join-Path $pythonSitePackages "nvidia\cudnn\bin"
        $cudnnDll = Join-Path $cudnnPath "cudnn_ops_infer64_8.dll"
        
        if (Test-Path $cudnnDll) {
            Write-InstallLog "cuDNN DLL verified: $cudnnDll" -Level "SUCCESS"
            return $true
        } else {
            Write-InstallLog "cuDNN DLL not found at expected location" -Level "ERROR"
            return $false
        }
    } catch {
        Write-InstallLog "Exception during cuDNN installation: $_" -Level "ERROR"
        return $false
    }
}

function Install-WhisperX {
    Write-InstallLog "Installing WhisperX..." -Level "INFO"
    try {
        pip install whisperx --upgrade
        if ($LASTEXITCODE -eq 0) {
            Write-InstallLog "WhisperX installation completed" -Level "SUCCESS"
            return $true
        } else {
            Write-InstallLog "WhisperX installation failed" -Level "ERROR"
            return $false
        }
    } catch {
        Write-InstallLog "Exception during WhisperX installation: $_" -Level "ERROR"
        return $false
    }
}

function Test-WhisperXInstallation {
    Write-InstallLog "Testing WhisperX installation..." -Level "INFO"
    try {
        # Get Python site-packages path
        $pythonSitePackages = python -c "import site; print(site.getsitepackages()[1])" 2>$null
        
        # Set up environment with cuDNN paths
        $cudnnPath = Join-Path $pythonSitePackages "nvidia\cudnn\bin"
        $cublasPath = Join-Path $pythonSitePackages "nvidia\cublas\bin"
        $nvrtcPath = Join-Path $pythonSitePackages "nvidia\cuda_nvrtc\bin"
        $env:PATH = "$cudnnPath;$cublasPath;$nvrtcPath;" + $env:PATH
        
        # Test WhisperX import
        $testResult = python -c "import whisperx; print('WhisperX import: SUCCESS')" 2>&1
        if ($testResult -match "SUCCESS") {
            Write-InstallLog "WhisperX import test passed" -Level "SUCCESS"
            
            # Test CUDA availability in WhisperX context
            $cudaTest = python -c "import torch; print(f'CUDA available: {torch.cuda.is_available()}'); print(f'GPU count: {torch.cuda.device_count()}')" 2>&1
            Write-InstallLog "CUDA test results: $cudaTest" -Level "INFO"
            return $true
        } else {
            Write-InstallLog "WhisperX import failed: $testResult" -Level "ERROR"
            return $false
        }
    } catch {
        Write-InstallLog "Exception during WhisperX testing: $_" -Level "ERROR"
        return $false
    }
}

function New-WhisperDirectories {
    Write-InstallLog "Creating WhisperX directory structure..." -Level "INFO"
    $directories = @(
        $InstallPath,
        "$InstallPath\watch",
        "$InstallPath\transcripts", 
        "$InstallPath\completed",
        "$InstallPath\failed",
        "$InstallPath\logs",
        "$InstallPath\service"
    )
    
    foreach ($dir in $directories) {
        if (-not (Test-Path $dir)) {
            try {
                New-Item -ItemType Directory -Path $dir -Force | Out-Null
                Write-InstallLog "Created directory: $dir" -Level "INFO"
            } catch {
                Write-InstallLog "Failed to create directory: $dir - $_" -Level "ERROR"
                return $false
            }
        } else {
            Write-InstallLog "Directory exists: $dir" -Level "INFO"
        }
    }
    return $true
}

function Get-PythonExecutablePath {
    try {
        $pythonPath = python -c "import sys; print(sys.executable)" 2>$null
        if ($pythonPath) {
            return $pythonPath.Trim()
        }
    } catch { }
    return $null
}

function Get-WhisperXExecutablePath {
    try {
        $pythonPath = Get-PythonExecutablePath
        if ($pythonPath) {
            $pythonDir = Split-Path $pythonPath
            $scriptsDir = Join-Path $pythonDir "Scripts"
            $whisperxPath = Join-Path $scriptsDir "whisperx.exe"
            if (Test-Path $whisperxPath) {
                return $whisperxPath
            }
        }
    } catch { }
    return $null
}

function Show-InstallationSummary {
    Write-InstallLog "=== WhisperX GPU Installation Summary ===" -Level "SUCCESS"
    
    # Python info
    $pythonPath = Get-PythonExecutablePath
    Write-InstallLog "Python executable: $pythonPath" -Level "INFO"
    
    # WhisperX info
    $whisperxPath = Get-WhisperXExecutablePath
    Write-InstallLog "WhisperX executable: $whisperxPath" -Level "INFO"
    
    # GPU info
    try {
        $gpuInfo = nvidia-smi --query-gpu=name,driver_version,memory.total --format=csv,noheader,nounits 2>$null
        Write-InstallLog "GPU: $gpuInfo" -Level "INFO"
    } catch { }
    
    # PyTorch info
    try {
        $torchInfo = python -c "import torch; print(f'PyTorch: {torch.__version__}, CUDA: {torch.cuda.is_available()}')" 2>$null
        Write-InstallLog "PyTorch: $torchInfo" -Level "INFO"
    } catch { }
    
    Write-InstallLog "Installation directory: $InstallPath" -Level "INFO"
    Write-InstallLog "Service script location: $InstallPath\service\whisper-service.ps1" -Level "INFO"
    Write-InstallLog "" -Level "INFO"
    Write-InstallLog "Next steps:" -Level "SUCCESS"
    Write-InstallLog "1. Place audio files in: $InstallPath\watch" -Level "INFO"
    Write-InstallLog "2. Run the service: powershell -File '$InstallPath\service\whisper-service.ps1'" -Level "INFO"
    Write-InstallLog "3. Check transcripts in: $InstallPath\transcripts" -Level "INFO"
    Write-InstallLog "4. Monitor logs in: $InstallPath\logs" -Level "INFO"
}

# =============================
# Main Installation Script
# =============================
function Start-WhisperXInstallation {
    Write-InstallLog "Starting WhisperX GPU installation..." -Level "SUCCESS"
    Write-InstallLog "Target installation path: $InstallPath" -Level "INFO"
    
    # System validation
    if (-not $SkipNvidiaCheck) {
        if (-not (Test-NvidiaGPU)) {
            Write-InstallLog "NVIDIA GPU validation failed. Use -SkipNvidiaCheck to bypass." -Level "ERROR"
            return $false
        }
    }
    
    if (-not (Test-PythonVersion)) {
        Write-InstallLog "Python validation failed. Please install Python 3.10 or newer." -Level "ERROR"
        return $false
    }
    
    # Create directories
    if (-not (New-WhisperDirectories)) {
        Write-InstallLog "Failed to create required directories" -Level "ERROR"
        return $false
    }
    
    # Install components
    Write-InstallLog "Installing required components..." -Level "INFO"
    
    if (-not (Install-PyTorchCUDA)) {
        Write-InstallLog "PyTorch CUDA installation failed" -Level "ERROR"
        return $false
    }
    
    if (-not (Install-CUDNNLibraries)) {
        Write-InstallLog "cuDNN installation failed" -Level "ERROR"
        return $false
    }
    
    if (-not (Install-WhisperX)) {
        Write-InstallLog "WhisperX installation failed" -Level "ERROR"
        return $false
    }
    
    # Test installation
    if (-not (Test-WhisperXInstallation)) {
        Write-InstallLog "WhisperX installation test failed" -Level "ERROR"
        return $false
    }
    
    # Success
    Write-InstallLog "WhisperX GPU installation completed successfully!" -Level "SUCCESS"
    Show-InstallationSummary
    return $true
}

# =============================
# Script Entry Point
# =============================
try {
    # Check if running as administrator
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    $isAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    
    if (-not $isAdmin) {
        Write-InstallLog "Note: Running without administrator privileges. Some operations may require elevation." -Level "WARN"
    }
    
    # Run installation
    $success = Start-WhisperXInstallation
    
    if ($success) {
        Write-InstallLog "Installation completed successfully!" -Level "SUCCESS"
        exit 0
    } else {
        Write-InstallLog "Installation failed. Check the errors above." -Level "ERROR"
        exit 1
    }
} catch {
    Write-InstallLog "Unexpected error during installation: $_" -Level "ERROR"
    exit 1
}
