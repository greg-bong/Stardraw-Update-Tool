# Stardraw Update Tool

Desktop tool for updating Stardraw model attributes from a source product export workbook.

## What It Does

- reads a source `.xlsx` product export
- matches rows by `Model Number`
- updates the destination Stardraw attributes workbook
- detects attribute conflicts before writing
- lets the user choose conflict resolutions in the GUI
- creates a backup copy of the live destination workbook before replacement

## Prerequisites For Operation

Before running the tool, make sure all of the following are true:

- Google Drive is running if the destination workbook or exclusion list file is stored in Google Drive
- the destination workbook is fully synced locally and opens normally in Excel
- the exclusion list file is fully synced locally if it is stored in Google Drive
- the destination workbook is not open in Excel by another user when the write step begins
- the source file is a valid `.xlsx` workbook
- the destination file is a valid `.xlsx` workbook
- the archive backup folder exists or can be created
- the user has write access to the destination workbook and the selected archive folder
- the user has read access to the exclusion list file
- the required fields in the app are selected:
  `Source Products Export`, `Destination Attributes File`, `Archive Backup Folder`, and `Exclusion List File`

Recommended operational setup:

- keep the exclusion list in a shared Google Drive folder
- keep the destination workbook in a shared Google Drive location that is available offline
- use an archive backup folder that all intended operators can access

## Main Files

- `app.py`: Tkinter GUI
- `engine.py`: workbook update pipeline
- `drive_check.py`: Google Drive preflight checks
- `device_id_exclusions.txt`: Device ID prefixes that should be cleared instead of normalized
- `app.spec`: PyInstaller packaging config
- `run.command`: macOS launcher that uses the local virtual environment

## Daily Use

### macOS

Fastest day-to-day option:

- open `dist/StardrawUpdateTool.app` directly
- or move `StardrawUpdateTool.app` into `Applications`
- optionally add it to the Dock for one-click launching

From Terminal:

```bash
cd /path/to/StardrawUpdateTool
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
- `Exclusion List File`

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

The exclusion list contains Device ID prefixes that should be cleared instead of normalized.

Recommended setup:

- store the exclusion list as a shared `.txt` file in Google Drive
- point the app's `Exclusion List File` field at that shared file
- let each user read the same central exclusions list

The app shows whether the selected exclusion file loaded successfully at startup and during runs.

## Google Drive Behavior

Google Drive checks only run when the destination path looks like a Google Drive location.

This means:

- real Google Drive destinations still get a Drive running check
- local test files on a Mac or Windows VM do not require Google Drive to be installed
- the exclusion list can also be loaded from a shared Google Drive text file

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

### Automated Builds (GitHub Actions)

The repository includes a GitHub Actions workflow that builds:

- a macOS app bundle zip
- a Windows app zip

Workflow file:

- `.github/workflows/build-packages.yml`

How to use it:

- open the repository in GitHub
- go to `Actions`
- run `Build Packages`
- download the uploaded artifacts from the workflow run

### GitHub Releases

The repository also includes a release workflow that creates downloadable GitHub Release assets when a version tag is pushed.

Workflow file:

- `.github/workflows/release-packages.yml`

Release flow:

1. push a tag like `v1.7.0`
2. GitHub Actions builds both packages
3. GitHub creates a Release for that tag
4. the Mac and Windows zip files are attached to the Release

This is the easiest download path for non-technical users because they can use the `Releases` page instead of opening workflow artifacts.

### Build macOS `.app`

From the project folder:

```bash
cd /path/to/StardrawUpdateTool
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
