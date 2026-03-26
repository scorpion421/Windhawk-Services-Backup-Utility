# Windhawk Service Management Utility

A lightweight Windows desktop tool for backing up and restoring [Windhawk](https://windhawk.net/) configurations, built with Python and tkinter.

**Please note:** There will be no compiled versions anymore!

![Python](https://img.shields.io/badge/Python-3.10%2B-blue)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey)
![Version](https://img.shields.io/badge/Version-2.5.0-green)
![License](https://img.shields.io/badge/License-GPL-yellow)

<img width="822" height="712" alt="image" src="https://github.com/user-attachments/assets/68776dfd-1299-49fb-85c5-bbb610f67941" />

---

## Features

- **One-click backup** of mod sources, compiled mods, and registry settings into a timestamped ZIP archive
- **One-click restore** from any previously created archive
- **Backup archive list** with sortable columns (date, size, type, mod count, filename)
- **Backup preview** — double-click any archive to inspect its contents (mod list, creation date, installation type, archive size) without extracting it
- **Portable installation support** — checkbox and auto-detect button skip registry export/import for portable Windhawk installs
- **Service management** — automatically stops the Windhawk service before any file operation and restarts it afterwards, even on error
- **Archive integrity check** — validates every new backup with `zipfile.testzip()` immediately after creation
- **Installation validation** — verifies the specified Windhawk root path before any operation
- **Backup rotation** — automatically removes the oldest archives when a configurable limit is exceeded (0 = unlimited)
- **Backup manifest** — each archive contains a `manifest.json` with creation timestamp, mod list, mod count, and installation type
- **Persistent settings** — paths, portable flag, and rotation limit are saved to `%AppData%\Windhawk_Backup_Utility\config.json` and restored on next launch
- **Timestamped operation log** with colour-coded entries (info, success, warning, error) and export to `.txt`
- **Status bar** showing current state at all times
- **Threaded operations** — UI remains fully responsive during backup and restore

---

## Requirements

- Windows 10 or 11
- Python 3.10 or newer (standard library only — no third-party packages required)
- Administrator privileges (required for registry access and service control)

---

## Installation

No installation needed. Clone or download the repository and run the script directly:

```bat
git clone https://github.com/scorpion421/wsbu.git
cd wsbu
python wsbu.py
```

If Python is not in your PATH, right-click `wsbu.py` and choose **Run as administrator**, or launch it from an elevated command prompt.

The tool will automatically request elevation via UAC if not already running as administrator.

---

## Usage

### Backup

1. Verify the **Windhawk Root** path (default: `%ProgramData%\Windhawk`)
2. Verify the **Backup Folder** path (default: `%UserProfile%\Documents\Windhawk_Backup`)
3. Set the desired number of backups to keep (0 = unlimited)
4. Click **Create Backup**

The tool will stop the Windhawk service, copy mod sources and compiled mods, export the registry key, compress everything into a ZIP archive, and restart the service.

### Restore

1. Select a backup from the archive list
2. Click **Restore Selected** (or double-click the entry for a preview first)
3. Confirm the dialog

The tool will stop the Windhawk service, extract the archive, restore all files, import the registry key, and restart the service.

### Portable installations

If Windhawk is running as a portable install (no registry key present):

- Tick **Portable installation** to skip all registry steps, or
- Click **Auto-Detect** to let the tool check automatically

### Preview

Double-click any archive in the list to open a detail window showing the full mod list, creation date, installation type, and archive size — without extracting the archive.

---

## Default paths

| Item | Path |
|------|------|
| Windhawk root | `%ProgramData%\Windhawk` |
| Backup folder | `%UserProfile%\Documents\Windhawk_Backup` |
| Settings file | `%AppData%\Windhawk_Backup_Utility\config.json` |
| Archive naming | `windhawk-backup_YYYYMMDD_HHMMSS.zip` |

---

## Changelog

### 2.5.0
- Added backup archive list with sortable columns and zebra striping
- Added backup preview dialog (double-click) showing full mod list and metadata
- Added `manifest.json` inside each archive (creation time, mod list, installation type)
- Added backup rotation — configurable maximum number of backups to keep
- Added persistent settings (paths, portable flag, rotation limit saved across sessions)
- Added timestamped log entries and log export to `.txt`
- Added status bar
- Added Windhawk service stop/start around all file operations (via `try/finally`)
- Added archive integrity check (`zipfile.testzip()`) after every backup
- Added Windhawk root validation before any operation
- Added portable installation checkbox and auto-detect button
- Operations now run on a background thread — UI no longer freezes
- Improved error handling and reporting throughout

### 2.1.1 (original)
- Initial release

---

This script is inspired by the work of https://github.com/lokize - https://github.com/ramensoftware/windhawk/issues/195#issuecomment-3184189085

---

## License

GPL License — see [LICENSE](LICENSE) for details.
