# where-file

[English](README.md) | [한국어](README.ko.md)

> "Where my files at?" — See every open file on your Windows PC, instantly.

No install. No dependencies. Just run it.

![Python](https://img.shields.io/badge/Python-3.6+-blue?logo=python&logoColor=white)
![Platform](https://img.shields.io/badge/Platform-Windows-0078D6?logo=windows&logoColor=white)
![Dependencies](https://img.shields.io/badge/Dependencies-Zero-brightgreen)
![License](https://img.shields.io/badge/License-MIT-yellow)

![where-file Web Dashboard](image/image_1.png)

## What is this?

A tiny Windows utility that shows you **all currently open files** across every app — PowerPoint, Word, Excel, VS Code, Notepad, and more.

Ever had 20 windows open and couldn't find that one file? **where-file** solves that.

### Key Features

- **See all open files** in one place (Office, editors, media players, etc.)
- **Copy file path** or **folder path** with one click
- **70+ file types** supported (Office, PDF, images, code, video, audio...)
- **Zero dependencies** — pure Python standard library only
- **Two modes**: Web dashboard or System tray

## Quick Start

![Getting Started](image/image_3.png)

```bash
# Clone
git clone https://github.com/kangyekwon/where-file.git
cd where-file

# Run (Web Dashboard)
python server.py
```

That's it. Browser opens automatically at `http://localhost:8765`.

## Two Ways to Use

### 1. Web Dashboard (`server.py`)

```bash
python server.py
```

- Beautiful dark-themed dashboard in your browser
- Filter files by type (Office, PDF, Image, Code, etc.)
- Auto-refresh every 5 seconds
- Copy path / Copy folder buttons

### 2. System Tray (`app.py`)

![System Tray Mode](image/image_2.png)

```bash
python app.py
```

- Sits quietly in your system tray
- **`Ctrl+Shift+F`** — instantly copies the active window's file path to clipboard
- Left-click tray icon to see all open files
- Perfect for quick "grab the path" moments

## Bonus: Right-Click Context Menu

Add "Copy Full Path" to your Windows Explorer right-click menu:

```
# Double-click to install
install_menu.reg

# Double-click to uninstall
uninstall_menu.reg
```

## Supported File Types

| Category | Extensions |
|----------|-----------|
| Office | `.pptx` `.docx` `.xlsx` `.hwp` `.hwpx` `.csv` |
| Documents | `.pdf` `.txt` `.rtf` `.md` `.log` |
| Images | `.png` `.jpg` `.gif` `.svg` `.webp` `.psd` `.ai` |
| Code | `.py` `.js` `.ts` `.java` `.cpp` `.go` `.rs` `.rb` `.php` |
| Video | `.mp4` `.avi` `.mkv` `.mov` |
| Audio | `.mp3` `.wav` `.flac` `.aac` |
| Archives | `.zip` `.rar` `.7z` `.tar` `.gz` |
| Design | `.sketch` `.fig` `.xd` `.indd` `.dwg` |
| And more... | 70+ extensions total |

## How It Works

1. **Office apps** (PowerPoint, Word, Excel): Queries COM objects directly via PowerShell
2. **Other apps**: Scans running process command lines and extracts file paths
3. Results are cached for 3 seconds to keep things fast

## Requirements

- **Windows** (uses Windows APIs)
- **Python 3.6+** (no pip install needed)
- **PowerShell 5.0+** (pre-installed on Windows 10/11)

## Project Structure

```
where-file/
├── server.py           # Web dashboard server
├── app.py              # System tray app with hotkey
├── index.html          # Web UI
├── image/              # Screenshots
├── install_menu.reg    # Explorer context menu (install)
└── uninstall_menu.reg  # Explorer context menu (uninstall)
```

## Known Limitations

- Windows only (relies on Windows APIs and PowerShell)
- Files opened in browser tabs (e.g., PDF in Chrome/Edge) are not detected
- Some apps may not expose file paths in their process command line

## License

MIT
