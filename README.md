# Stardraw Update Tool

Desktop tool for updating Stardraw model attributes from a source product export workbook.

## What It Does

- reads a source `.xlsx` product export
- matches rows by `Model Number`
- updates the destination Stardraw attributes workbook
- detects attribute conflicts before writing
- lets the user choose conflict resolutions in the GUI
- creates a backup copy of the live destination workbook before replacement

## Main Files

- `app.py`: Tkinter GUI
- `engine.py`: workbook update pipeline
- `drive_check.py`: Google Drive preflight checks
- `device_id_exclusions.txt`: Device ID prefixes that should be cleared instead of normalized
- `app.spec`: PyInstaller packaging config
- `run.command`: macOS launcher that uses the local virtual environment

## Daily Use

### macOS

From Terminal:

```bash
cd /Users/gregoryknight/Desktop/StardrawUpdateTool_v1_4_1
./run.command
```

Or double-click `run.command` in Finder.

### Windows

Run `StardrawUpdateTool.exe` from inside the built Windows app folder.

Important:

- keep the whole `StardrawUpdateTool` folder together
- do not separate `StardrawUpdateTool.exe` from `_internal`

## Required Inputs In The App

The GUI requires:

- `Source Products Export`
- `Destination Attributes File`
- `Archive Backup Folder`

The tool will not run unless all three are selected.

## Backup Behavior

Before the live destination workbook is replaced, the tool creates a backup copy in the selected archive folder.

The backup is:

- a copy of the current live destination workbook
- created immediately before replacement
- named like `ARCHIVE_<timestamp>_<destination filename>.xlsx`

This is the restore point if something goes wrong after the backup step.

## Conflict Resolution

If multiple values are found for the same model/field combination:

- the run stops before writing
- a conflict chooser window opens
- the user selects one value for each conflict
- the tool reruns using those selected values

## Device ID Exclusions

The file `device_id_exclusions.txt` contains Device ID prefixes that should be cleared instead of normalized.

The app shows whether this file loaded successfully at startup.

## Google Drive Behavior

Google Drive checks only run when the destination path looks like a Google Drive location.

This means:

- real Google Drive destinations still get a Drive running check
- local test files on a Mac or Windows VM do not require Google Drive to be installed

## Local Development Setup

### macOS

Create and use the local virtual environment:

```bash
python3 -m venv venv
./venv/bin/pip install pandas openpyxl pyinstaller pillow
```

Run the app:

```bash
./venv/bin/python app.py
```

### Windows

Use a real local Windows folder, not a redirected `\\Mac\...` path.

Example:

```powershell
cd "C:\Users\<your-user>\Desktop\Stardraw-Update-Tool\Stardraw-Update-Tool"
..\venv\Scripts\python -m pip install pandas openpyxl pyinstaller
..\venv\Scripts\python app.py
```

## Packaging

### Build macOS `.app`

From the project folder:

```bash
cd /Users/gregoryknight/Desktop/StardrawUpdateTool_v1_4_1
env PYINSTALLER_CONFIG_DIR="$PWD/.pyinstaller" ./venv/bin/pyinstaller --noconfirm --distpath "$PWD/dist" --workpath "$PWD/build" app.spec
```

Main output:

- `dist/StardrawUpdateTool.app`

Optional zip:

```bash
ditto -c -k --keepParent dist/StardrawUpdateTool.app dist/StardrawUpdateTool-macOS.zip
```

### Build Windows `.exe`

From the inner repo folder in the Windows VM:

```powershell
..\venv\Scripts\python -m PyInstaller --noconfirm --distpath ..\dist --workpath ..\build app.spec
```

Main output:

- `dist\StardrawUpdateTool\StardrawUpdateTool.exe`

Important:

- this is a `onedir` build
- distribute the whole `dist\StardrawUpdateTool` folder

## Windows Build Notes

- Windows Explorer may hide the `.exe` extension
- if Explorer shows `StardrawUpdateTool` as a file inside the `dist\StardrawUpdateTool` folder, that is usually `StardrawUpdateTool.exe`
- the project includes a Windows icon file: `Twisted_Icon.ico`

## Recommended Distribution

### macOS

- share `StardrawUpdateTool.app`
- or share `StardrawUpdateTool-macOS.zip`

### Windows

- zip the whole `dist\StardrawUpdateTool` folder
- users extract it locally
- users launch `StardrawUpdateTool.exe`

## Troubleshooting

### `Google Drive not running`

The destination path looks like Google Drive, but Drive is not running on that machine.

### `Destination file is not a valid local .xlsx workbook`

The selected destination is not a real local Excel workbook at runtime.

Common causes:

- shortcut instead of real file
- cloud-only placeholder
- invalid workbook file

### `Attribute conflicts detected`

The source workbook contains multiple values for the same model/field combination.

Use the conflict chooser to select the value that should win.

### Windows PowerShell cannot find `pyinstaller`

Use the venv Python directly:

```powershell
..\venv\Scripts\python -m PyInstaller --noconfirm --distpath ..\dist --workpath ..\build app.spec
```

## Repository

GitHub:

- [greg-bong/Stardraw-Update-Tool](https://github.com/greg-bong/Stardraw-Update-Tool)
