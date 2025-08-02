# Verxell

**A powerful Excel processing tool for duplicating rows with headers**

Verxell is a desktop application that automates the process of restructuring Excel files by duplicating data rows with their corresponding headers. It's particularly useful for data formatting, report generation, and Excel file manipulation tasks.

## ✨ Features

- **Drag & Drop Interface**: Simply drag Excel files into the application
- **Batch Processing**: Process multiple Excel files simultaneously
- **Header Detection**: Automatically detects colored headers (yellow by default)
- **Row Duplication**: Creates duplicates of data rows with proper formatting
- **Formula Preservation**: Maintains and adjusts Excel formulas during duplication
- **Multi-language Support**: Available in English and Russian
- **Progress Tracking**: Real-time progress monitoring with pause/resume functionality
- **Error Reporting**: Built-in error reporting system via Telegram
- **Auto-Updates**: Automatic update checking and installation
- **Comprehensive Logging**: Detailed logging for troubleshooting

## 📦 Installation

### Option 1: Download Pre-built Executable
1. Go to [Releases](https://github.com/Slipfaith/DM/releases)
2. Download the latest `Verxell.exe`
3. Run the executable directly (no installation required)

### Option 2: Run from Source
1. Clone the repository:
```bash
git clone https://github.com/Slipfaith/DM.git
cd DM
```

2. Install Python dependencies:
pip install -r requirements.txt


3. Run the application:
python main.py


## 🚀 Quick Start

1. **Launch Verxell**
2. **Add Files**: 
   - Drag and drop Excel files onto the interface, or
   - Click the drop area to browse and select files
3. **Process Files**: Click "Process Files" to start
4. **Monitor Progress**: Watch real-time progress and logs
5. **Access Results**: Processed files are saved in a "Deeva" folder next to originals



### Python Requirements (if running from source)
- Python 3.8+
- PySide6 >= 6.5.0
- pywin32 >= 305
- openpyxl >= 3.1.0
- packaging >= 23.0

## 🔧 How It Works

Verxell processes Excel files by:

1. **Scanning for Headers**: Identifies header rows by color (yellow/65535 by default)
2. **Locating Data Blocks**: Finds data rows following each header
3. **Duplicating Rows**: Creates exact duplicates of data rows
4. **Adjusting Formulas**: Updates cell references in formulas (especially LEN/ДЛСТР functions)
5. **Adding Structure**: Inserts spacing and additional headers for clarity
6. **Preserving Formatting**: Maintains original cell formatting and styles

### Example Transformation

**Before:**

Header Row (Yellow)
Data Row 1
Data Row 2
```

**After:**

Header Row (Yellow)
Data Row 1
Data Row 1 (Duplicate)
[Empty Row]
Header Row (Yellow)
Data Row 2
Data Row 2 (Duplicate)
[Empty Row]
```

## 🌐 Internationalization

Verxell supports multiple languages:
- **English** (default)
- **Russian** (Русский)

Change language through the menu: `Language → English/Русский`

## 📝 Logging

Comprehensive logging is available in the `logs/` directory:
- Timestamped log files for each session
- Real-time log display in the application
- Error tracking and debugging information

## 🐛 Error Reporting

Built-in error reporting system:
- Automatic error detection
- Optional error reporting to developer via Telegram
- Attach screenshots and log files
- Privacy-conscious (only sends what you approve)

## 🔄 Updates

Verxell includes automatic update functionality:
- Checks for updates on startup
- Manual update checking via Help menu
- Secure downloads with hash verification
- Automatic installation (for executable versions)

## 🛠️ Development

### Project Structure
```
verxell/
├── main.py              # Application entry point
├── gui.py               # Main GUI interface
├── excel_processor.py   # Core Excel processing logic
├── excel_com.py         # Excel COM interface wrapper
├── config.py            # Configuration management
├── logger.py            # Logging utilities
├── translations.py      # Internationalization
├── updater.py           # Auto-update functionality
├── error_dialog.py      # Error reporting dialogs
├── settings_manager.py  # Settings persistence
├── styles.py            # UI styling
├── telegram/            # Telegram integration
│   ├── __init__.py
│   ├── reporter.py      # Error reporting
│   └── config.py        # Telegram configuration
└── requirements.txt     # Python dependencies
```

### Building Executable
To create a standalone executable:

1. Install PyInstaller:
```bash
pip install pyinstaller
```

2. Build the executable:
```bash
pyinstaller --onefile --windowed --icon=icon.ico main.py
```

### Contributing
1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## 📄 License

This project is released under the MIT License. See [LICENSE](LICENSE) file for details.

## 🤝 Support

### Getting Help
- Check the [Issues](https://github.com/Slipfaith/DM/issues) page for known problems
- Use the built-in "Write to Owner" feature for direct support
- Create a new issue for bug reports or feature requests

### Reporting Bugs
When reporting bugs, please include:
- Operating system and version
- Excel version
- Steps to reproduce the issue
- Sample Excel file (if possible)
- Error messages or logs

## 🏷️ Version History

### v1.2.0 (Current)
- Added multi-language support
- Improved error reporting
- Enhanced UI/UX
- Better formula handling
- Auto-update functionality

### v1.1.0
- Added drag and drop support
- Improved progress tracking
- Better error handling
- Added logging system

### v1.0.0
- Initial release
- Basic Excel processing
- Simple GUI interface

## 🔗 Links

- **Repository**: https://github.com/Slipfaith/DM
- **Issues**: https://github.com/Slipfaith/DM/issues
- **Releases**: https://github.com/Slipfaith/DM/releases
<img width="791" height="624" alt="update" src="https://github.com/user-attachments/assets/5ac9d08f-d384-4434-88e1-a3b5b59bfdde" />
<img width="396" height="323" alt="Initial-window" src="https://github.com/user-attachments/assets/ebec24a2-5cf1-4f8a-a903-42f7f5cd79b3" />
<img width="793" height="627" alt="Feedback" src="https://github.com/user-attachments/assets/931332c1-e171-4219-9673-5bddd5158362" />
<img width="794" height="624" alt="working" src="https://github.com/user-attachments/assets/222ae1ab-4f95-45b2-a73b-767cb0895fc7" />

