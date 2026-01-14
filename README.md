# WhisperX GPU Auto-Transcription Service

A Windows PowerShell service that automatically transcribes audio files using OpenAI's WhisperX with NVIDIA GPU acceleration. This service watches a directory for new audio files and processes them using your RTX GPU for fast, accurate transcription.

## 🚀 Features

- **GPU Acceleration**: Utilizes NVIDIA RTX GPUs for fast transcription
- **Automatic Processing**: Watches directory for new audio files
- **Multiple Formats**: Supports M4A, MP3, WAV, FLAC, AAC, OGG, WMA, M4B, WebM
- **File Organization**: Automatically moves completed/failed files
- **Comprehensive Logging**: Detailed logs for monitoring and troubleshooting
- **Space-Safe Processing**: Handles filenames with spaces correctly
- **Robust Error Handling**: Graceful shutdown and error recovery

## 📋 System Requirements

### Hardware
- **GPU**: NVIDIA RTX 2000 series or newer (RTX 3070, 4070, etc.)
- **VRAM**: Minimum 4GB, 8GB+ recommended
- **RAM**: 8GB+ system RAM
- **Storage**: 2GB+ free space for models and cache

### Software
- **OS**: Windows 10/11 (64-bit)
- **NVIDIA Drivers**: Latest drivers installed
- **Python**: 3.10 or newer
- **PowerShell**: 5.1 or newer

## 🔧 Quick Installation

### Option 1: Automated Installation (Recommended)

1. **Download the installer script** to your desired location (e.g., `C:\whisper\`)

2. **Run the installation script** as Administrator:
   ```powershell
   powershell -ExecutionPolicy Bypass -File "C:\whisper\install-whisperx-gpu.ps1"
   ```

3. **Wait for completion** - The script will:
   - Validate your system
   - Install PyTorch with CUDA support
   - Install cuDNN libraries
   - Install WhisperX
   - Create directory structure
   - Test the installation

### Option 2: Custom Installation Path

```powershell
powershell -ExecutionPolicy Bypass -File "C:\whisper\install-whisperx-gpu.ps1" -InstallPath "D:\MyWhisper"
```

### Option 3: Manual Installation

If you prefer manual installation:

1. **Install PyTorch with CUDA 11.8**:
   ```bash
   pip install torch torchvision torchaudio --index-url https://download.pytorch.org/whl/cu118
   ```

2. **Install cuDNN 8.x**:
   ```bash
   pip install nvidia-cudnn-cu11==8.9.4.25
   ```

3. **Install WhisperX**:
   ```bash
   pip install whisperx
   ```

4. **Create directory structure**:
   ```
   C:\whisper\
   ├── watch\          # Place audio files here
   ├── transcripts\    # Completed transcriptions
   ├── completed\      # Successfully processed audio files
   ├── failed\         # Failed audio files
   ├── logs\           # Service logs
   └── service\        # Service script
   ```

## 🎯 How It Works

### Directory Structure
```
C:\whisper\
├── watch\                    # 📁 Input: Place audio files here for processing
├── transcripts\              # 📄 Output: Text transcriptions appear here
├── completed\                # ✅ Archive: Successfully processed audio files
├── failed\                   # ❌ Archive: Files that failed processing
├── logs\                     # 📋 Logs: Daily service logs
├── service\
│   └── whisper-service.ps1   # 🔧 Main service script
├── install-whisperx-gpu.ps1  # 📦 Installation script
└── README.md                 # 📖 This file
```

### Processing Workflow

1. **File Detection**: Service monitors `watch\` directory for new audio files
2. **GPU Processing**: WhisperX processes files using your RTX GPU with:
   - Model: Medium (good balance of speed/accuracy)
   - Device: CUDA (GPU acceleration)
   - Compute Type: float16 (optimal for RTX GPUs)
3. **File Management**: 
   - ✅ **Success**: Audio → `completed\`, Transcript → `transcripts\`
   - ❌ **Failure**: Audio → `failed\`, Error logged
4. **Logging**: All operations logged to `logs\whisper-service-YYYYMMDD.log`

### Supported Audio Formats

| Format | Extension | Notes |
|--------|-----------|-------|
| M4A    | .m4a      | Apple audio format |
| MP3    | .mp3      | Common compressed format |
| WAV    | .wav      | Uncompressed audio |
| FLAC   | .flac     | Lossless compression |
| AAC    | .aac      | Advanced audio codec |
| OGG    | .ogg      | Open source format |
| WMA    | .wma      | Windows media audio |
| M4B    | .m4b      | Audiobook format |
| WebM   | .webm     | Web media format |

## 🚀 Usage

### Starting the Service

1. **Run the service**:
   ```powershell
   powershell -File "C:\whisper\service\whisper-service.ps1"
   ```

2. **Place audio files** in `C:\whisper\watch\`

3. **Monitor progress**:
   - Check logs in `C:\whisper\logs\`
   - Watch for transcripts in `C:\whisper\transcripts\`
   - Processed files move to `C:\whisper\completed\`

### Service Controls

- **Stop Service**: Press `Ctrl+C` in the PowerShell window
- **View Logs**: Check `C:\whisper\logs\whisper-service-YYYYMMDD.log`
- **Monitor GPU**: Use `nvidia-smi` to see GPU usage during processing

### Example Usage

1. Copy `meeting_recording.m4a` to `C:\whisper\watch\`
2. Service detects the file and starts processing
3. GPU processes the audio using WhisperX
4. Transcript appears as `meeting_recording.txt` in `C:\whisper\transcripts\`
5. Original audio moves to `C:\whisper\completed\meeting_recording.m4a`

## ⚙️ Configuration

### Service Configuration

Edit the configuration section in `whisper-service.ps1`:

```powershell
[string]$WatchPath     = "C:\whisper\watch"      # Input directory
[string]$OutputPath    = "C:\whisper\transcripts" # Output directory  
[string]$CompletedPath = "C:\whisper\completed"   # Completed files
[string]$FailedPath    = "C:\whisper\failed"      # Failed files
[string]$LogPath       = "C:\whisper\logs"        # Log directory
[string]$Language      = "en"                     # Language code
[int]$CheckInterval    = 5                        # Check interval (seconds)
```

### WhisperX Configuration

The service uses these optimal settings for RTX GPUs:

```powershell
$whisperxArgs = @(
    "--model", "medium",           # Model size (tiny, base, small, medium, large)
    "--language", $Language,       # Language code (en, es, fr, etc.)
    "--output_format", "txt",      # Output format (txt, json, srt, vtt)
    "--output_dir", $OutputPath,   # Output directory
    "--device", "cuda",            # Use GPU acceleration
    "--compute_type", "float16",   # Optimal for RTX GPUs
    $quotedFilePath               # Input file path
)
```

### Model Options

| Model | Size | VRAM | Speed | Accuracy |
|-------|------|------|-------|----------|
| tiny  | 39MB | <1GB | Fastest | Basic |
| base  | 74MB | <1GB | Fast | Good |
| small | 244MB | ~2GB | Medium | Better |
| medium| 769MB | ~4GB | Slower | Great |
| large | 1550MB| ~8GB | Slowest | Best |

## 🔧 Troubleshooting

### Common Issues

#### GPU Not Detected
```
[ERROR] WhisperX CLI executable not found...
```
**Solution**: Run the installation script to ensure proper GPU setup.

#### cuDNN DLL Missing
```
Could not locate cudnn_ops_infer64_8.dll
```
**Solution**: The installation script installs the correct cuDNN version (8.9.4.25).

#### Out of Memory
```
CUDA out of memory
```
**Solutions**:
- Use a smaller model (`small` instead of `medium`)
- Close other GPU-intensive applications
- Process shorter audio files

#### File Path Issues
```
Error opening input file
```
**Solution**: The service handles spaces in filenames automatically with proper quoting.

### Performance Optimization

#### For RTX 3070 (8GB VRAM):
- **Recommended Model**: `medium` (best balance)
- **Max Model**: `large` (may hit VRAM limits with long files)
- **Expected Speed**: ~10x faster than CPU

#### For RTX 4070+ (12GB+ VRAM):
- **Recommended Model**: `large` (best accuracy)
- **Expected Speed**: ~15x faster than CPU

### Log Analysis

Check daily logs in `C:\whisper\logs\`:

```
[2025-09-15 13:00:48] [INFO] Starting transcription for: voice_memo.m4a
[2025-09-15 13:00:48] [INFO] Using whisperx from full path: ...whisperx.exe
[2025-09-15 13:00:48] [INFO] Running whisperx command: ...
[2025-09-15 13:01:30] [SUCCESS] Successfully processed: voice_memo.m4a
[2025-09-15 13:01:30] [INFO] Transcript file created: voice_memo.txt
[2025-09-15 13:01:30] [INFO] Moved to completed: voice_memo.m4a
```

## 🔄 Updates and Maintenance

### Updating WhisperX
```powershell
pip install whisperx --upgrade
```

### Updating PyTorch
```powershell
pip install torch torchvision torchaudio --index-url https://download.pytorch.org/whl/cu118 --upgrade
```

### Log Rotation
Logs are created daily. Old logs can be safely deleted to save space.

## 🏗️ Technical Details

### Dependencies
- **PyTorch**: 2.7.1+cu118 (CUDA 11.8 support)
- **cuDNN**: 8.9.4.25 (CUDA Deep Neural Network library)  
- **WhisperX**: Latest (enhanced Whisper with speaker diarization)
- **faster-whisper**: Core transcription engine
- **ctranslate2**: Optimized inference library

### GPU Memory Usage
- **Model Loading**: ~1-2GB (one-time)
- **Audio Processing**: ~2-4GB (depends on file length)
- **Peak Usage**: Up to 6GB for large models with long files

### Processing Speed (RTX 3070)
- **Real-time Factor**: ~0.1 (10x faster than real-time)
- **Example**: 1-hour audio processes in ~6 minutes
- **Batch Processing**: Multiple files queued automatically

## 📝 License and Credits

This project uses:
- **WhisperX**: Enhanced version of OpenAI's Whisper
- **OpenAI Whisper**: Original speech recognition model
- **PyTorch**: Deep learning framework
- **NVIDIA CUDA**: GPU acceleration platform

## 🤝 Contributing

Feel free to submit issues, feature requests, or improvements to this transcription service.

---

**Enjoy fast, accurate transcriptions with your NVIDIA RTX GPU! 🚀**
