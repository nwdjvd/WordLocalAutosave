# Word Local Autosave

![Python Application](https://github.com/nwdjvd/WordLocalAutosave/workflows/Python%20Application/badge.svg)
![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)
![Python 3.8+](https://img.shields.io/badge/python-3.8+-blue.svg)

A robust offline autosave utility for Microsoft Word that works similarly to OneDrive's autosave functionality, but saves locally. It automatically monitors and saves your Word documents when changes are detected.

## Features

- **Event-based Autosave**: Listens for Word events to save documents when changes occur
- **Polling Backup**: Additional polling system to catch any changes that might be missed
- **Debounce Logic**: Prevents excessive saving (configurable, default: 10 seconds)
- **Multiple Document Support**: Works with all open Word documents simultaneously
- **Configurable**: Settings can be adjusted via JSON configuration file
- **Resilient Connection**: Automatically reconnects if Word is closed and reopened
- **Fault Tolerant**: Falls back to polling-only mode if event setup fails
- **Local Only**: Everything stays on your machine - no cloud or internet required

## Requirements

- Windows 10 or later
- Microsoft Word (desktop version)
- Python 3.8+ 
- PyWin32 package

## Installation

### From GitHub

1. **Clone or download this repository**:
   ```
   git clone https://github.com/nwdjvd/WordLocalAutosave.git
   cd WordLocalAutosave
   ```

2. **Install dependencies**:
   ```
   pip install -r requirements.txt
   ```

3. **Test your Word connection**:
   ```
   python test_word_connection.py
   ```
   This will verify if Word is correctly installed and can be accessed via COM.

### Using Pip (Coming Soon)

```
pip install word-local-autosave
```

## Usage

1. **Start the autosave service**:
   ```
   python main.py
   ```
   Or if installed via pip:
   ```
   word-autosave
   ```

2. **Work in Word normally**:
   - Open and edit documents in Word
   - The service will automatically save your documents when changes are made
   - Each document will be saved at most once every 10 seconds by default
   - The console and log file will show when documents are saved

3. **Stop the service**:
   - Press `Ctrl+C` in the console window to stop the service

## Configuration

You can customize the behavior by editing `autosave_config.json`:

```json
{
    "debounce_seconds": 10,     // Minimum time between saves for a document
    "polling_interval": 15,     // Seconds between backup polling checks
    "reconnect_threshold": 20,  // Number of attempts before reconnecting
    "main_loop_sleep": 0.5,     // Seconds to sleep in main loop
    "log_max_size": 5242880,    // 5 MB max log size
    "log_backup_count": 3,      // Number of log backups to keep
    "log_directory": ""         // Log directory (empty = current directory)
}
```

## Troubleshooting

If you encounter issues:

1. **Check the logs**:
   - The app creates `autosave.log` in the application directory (or configured directory)
   - Review error messages for clues about what's wrong

2. **Verify Word connection**:
   - Run `test_word_connection.py` to check your Word installation
   - Make sure Word is properly installed and can be accessed by Python

3. **COM issues**:
   - Try running the script as Administrator
   - Make sure your Word version is compatible (tested with Office 2016+)
   - Restart Word if you're experiencing connection problems

4. **No events firing**:
   - The app will automatically fall back to polling mode (checking documents every 5 seconds)
   - This is normal in some Word configurations and doesn't affect functionality

## How It Works

- Connects to Word using the Windows COM interface via pywin32
- Creates a custom event sink to listen for document changes
- Checks if documents are "dirty" (have unsaved changes) using the .Saved property
- Implements a debounce mechanism using monotonic time to avoid saving too frequently
- Uses a backup polling system that checks all open documents periodically
- Creates robust logging with automatic rotation to prevent excessive disk usage

## Limitations

- Works only with the desktop version of Word (not web/online versions)
- COM interface can sometimes be finicky depending on Word version and settings
- Event handlers might not catch every single change, but the polling backup ensures saves
- Administrator privileges may be required on some systems
- Does not save documents that have never been saved before (no path)

## Contributing

Contributions are welcome! Please see [CONTRIBUTING.md](CONTRIBUTING.md) for details on how to contribute to this project.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments

- Microsoft for the Word COM interface documentation
- PyWin32 project for making COM interfaces accessible in Python

## Future Improvements

- System tray application with status indicator
- Save version history (.bak files) 
- Manual save shortcut
- More detailed logging and statistics
- GUI for configuration
- Auto-recovery for crashed documents