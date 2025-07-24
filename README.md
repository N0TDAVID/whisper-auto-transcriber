# Whisper Auto-Transcriber Windows Service

A complete Windows service solution for automated audio transcription using OpenAI's Whisper technology. This PowerShell-based service monitors a directory for new audio files and automatically transcribes them using the Whisper CLI with GPU acceleration.

## 🚀 Features

- **Multi-Format Audio Support**: Automatically processes 9 audio formats (M4A, MP3, WAV, FLAC, AAC, OGG, WMA, M4B, WebM)
- **Automated File Monitoring**: Watches for new audio files in a designated directory
- **Sequential Processing**: Reliable one-file-at-a-time processing with comprehensive error handling
- **OpenAI Whisper Integration**: Uses locally installed Whisper CLI for high-quality transcription
- **GPU Acceleration**: Leverages your GPU for faster transcription processing
- **Windows Service**: Runs as a proper Windows service using NSSM
- **Health Monitoring**: Automatic service health checks and component restart
- **Daily Log Rotation**: Automatic log file management with detailed logging
- **Graceful Shutdown**: Proper cleanup and resource management
- **Startup File Processing**: Processes existing files when service starts

## 📋 Prerequisites

- Windows 10/11
- PowerShell 5.1 or higher
- Administrator privileges for installation
- NVIDIA GPU (recommended for faster processing)
- Python 3.10+ with Whisper CLI installed

## 🛠️ Installation

### Quick Install (Recommended)

1. **Download the project files** to your local machine
2. **Open PowerShell as Administrator**
3. **Navigate to the project directory**
4. **Run the installation script:**

```powershell
.\install.ps1
```

The installation script will automatically:
- ✅ Check for administrator privileges
- ✅ Download and install Python (if missing)
- ✅ Install Whisper CLI via pip
- ✅ Download and install ffmpeg (if missing)
- ✅ Download and install NSSM (if missing)
- ✅ Create required directory structure
- ✅ Configure and register the Windows service
- ✅ Set up logging and monitoring

### Manual Installation

If you prefer to install dependencies manually:

1. **Install Python 3.10+** from [python.org](https://www.python.org/downloads/)
2. **Install Whisper CLI:**
   ```powershell
   pip install git+https://github.com/openai/whisper.git
   ```
3. **Install ffmpeg** from [ffmpeg.org](https://ffmpeg.org/download.html)
4. **Download NSSM** from [nssm.cc](https://nssm.cc/download)
5. **Run the installation script** as above

## 📁 Directory Structure

After installation, the service creates the following structure:

```
C:\whisper\
├── watch\          # Place audio files here for transcription
├── transcripts\    # Generated transcript files (.txt)
├── completed\      # Successfully processed audio files
├── failed\         # Audio files that failed to process
├── logs\          # Daily log files
└── service\       # Service files
```

## 🔧 Configuration

### Service Configuration

The service uses these default settings (configurable in `whisper-service.ps1`):

- **Watch Directory**: `C:\whisper\watch`
- **Output Directory**: `C:\whisper\transcripts`
- **Completed Directory**: `C:\whisper\completed`
- **Failed Directory**: `C:\whisper\failed`
- **Log Directory**: `C:\whisper\logs`
- **Language**: `en` (English)
- **Check Interval**: `30` seconds
- **Whisper Model**: `medium` (good balance of speed and accuracy)

### Supported Audio Formats

The service automatically detects and processes these audio formats:
- **M4A** (AAC)
- **MP3**
- **WAV**
- **FLAC**
- **AAC**
- **OGG**
- **WMA**
- **M4B** (Audiobook)
- **WebM**

### Customizing Settings

Edit `whisper-service.ps1` to modify these parameters at the top of the file:

```powershell
[string]$WatchPath     = "C:\whisper\watch"
[string]$OutputPath    = "C:\whisper\transcripts"
[string]$CompletedPath = "C:\whisper\completed"
[string]$FailedPath    = "C:\whisper\failed"
[string]$LogPath       = "C:\whisper\logs"
[string]$Language      = "en"
[int]$CheckInterval    = 30  # seconds
```

## 🎯 Usage

### Starting the Service

After installation, the service will start automatically. To manually control it:

```powershell
# Start the service
Start-Service WhisperTranscriber

# Stop the service
Stop-Service WhisperTranscriber

# Check service status
Get-Service WhisperTranscriber
```

### Using the Service

1. **Place audio files** (any supported format) in `C:\whisper\watch`
2. **Wait up to 30 seconds** for file processing to begin
3. **Check `C:\whisper\transcripts`** for generated transcript files
4. **Successfully processed files** are moved to `C:\whisper\completed`
5. **Failed files** are moved to `C:\whisper\failed`

### File Processing

- **Sequential Processing**: Files are processed one at a time for reliability
- **Automatic Detection**: New files are detected within 30 seconds
- **Error Handling**: Failed files are moved to the failed directory with error logs
- **GPU Acceleration**: Transcription uses your GPU for faster processing

### Monitoring

- **Service Logs**: Check `C:\whisper\logs\whisper-service-YYYYMMDD.log`
- **Service Status**: Use `Get-Service WhisperTranscriber`
- **Event Viewer**: Check Windows Event Viewer for service events
- **GPU Usage**: Monitor GPU utilization during transcription

## 🗑️ Uninstallation

To remove the service and optionally clean up files:

```powershell
# Basic uninstall (keeps files)
.\uninstall.ps1

# Uninstall with directory cleanup
.\uninstall.ps1 -RemoveDirs
```

## 🔍 Troubleshooting

### Common Issues

**Service won't start:**
- Check administrator privileges
- Verify Whisper CLI is installed: `whisper --help`
- Check log files in `C:\whisper\logs`
- Ensure Python path is correct in the service script

**Files not being processed:**
- Ensure files are in a supported audio format
- Check file permissions
- Verify service is running: `Get-Service WhisperTranscriber`
- Check for files in the failed directory

**Transcription errors:**
- Check ffmpeg installation: `ffmpeg -version`
- Verify audio file integrity
- Check available disk space
- Ensure GPU drivers are up to date (for GPU acceleration)

**GPU not being used:**
- Verify CUDA is installed (for NVIDIA GPUs)
- Check GPU drivers are current
- Monitor GPU usage during transcription

### Log Analysis

The service creates detailed logs with timestamps and log levels:

```
[2024-01-15 14:30:25] [INFO] Service started successfully
[2024-01-15 14:30:30] [INFO] Processing file: audio.mp3
[2024-01-15 14:32:15] [INFO] Successfully processed: audio.mp3
[2024-01-15 14:32:15] [INFO] Moved to completed: audio.mp3
```

## 🏗️ Architecture

### Core Components

1. **FileSystemWatcher**: Monitors directory for new audio files
2. **Main Processing Loop**: Sequential file processing every 30 seconds
3. **Whisper Integration**: CLI execution with comprehensive error handling
4. **File Management**: Automated file movement and organization
5. **Health Monitoring**: Service health checks and component monitoring

### Service Lifecycle

1. **Startup**: Configuration validation → Directory creation → Component initialization
2. **Runtime**: File monitoring → Sequential processing → Health monitoring
3. **Shutdown**: Graceful cleanup → Resource release → Log finalization

## 📊 Performance

- **Processing Time**: Varies by audio length and GPU availability (typically 1-10 minutes per file)
- **Memory Usage**: ~100-200MB during operation
- **GPU Usage**: High during transcription (when GPU acceleration is available)
- **Disk I/O**: Low for monitoring, high during transcription
- **Multi-Format Support**: All supported audio formats processed with same efficiency

## 🔒 Security

- **Administrator Privileges**: Required for installation and service management
- **File Permissions**: Service runs with appropriate file access rights
- **Logging**: No sensitive data logged, only operational information
- **Local Processing**: All transcription happens locally, no data sent to external services

## 🤝 Contributing

This project is complete and production-ready. For issues or improvements:

1. Check the troubleshooting section
2. Review log files for detailed error information
3. Ensure all prerequisites are properly installed
4. Verify audio file formats are supported

## 📄 License

This project is provided as-is for educational and practical use.

## 🎉 Project Status

**✅ COMPLETE** - Production-ready with multi-format audio support

- ✅ Core PowerShell service structure
- ✅ Multi-format audio file monitoring
- ✅ Sequential file processing with error handling
- ✅ Whisper CLI integration with GPU support
- ✅ Automated file management and organization
- ✅ Service installation/uninstallation
- ✅ Comprehensive logging and monitoring
- ✅ Health monitoring and recovery
- ✅ Startup file processing
- ✅ Support for 9 audio formats

---

**Ready for production use!** 🚀 