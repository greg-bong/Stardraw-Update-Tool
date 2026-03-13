import os
import platform
import subprocess

GOOGLE_DRIVE_HINTS = (
    "google drive",
    "googledrive",
    "shared drives",
    "cloudstorage",
)


def path_looks_like_google_drive(folder_path):
    """Return True when the destination path appears to be a Google Drive location."""
    normalized = folder_path.replace("\\", "/").lower()
    return any(hint in normalized for hint in GOOGLE_DRIVE_HINTS)


def check_drive(folder_path):
    """Confirm the destination folder exists and Google Drive appears to be running."""
    if not os.path.exists(folder_path):
        raise Exception("Destination folder not found.")

    if not path_looks_like_google_drive(folder_path):
        return True

    system = platform.system()

    if system == "Darwin":  # macOS
        result = subprocess.run(["pgrep", "-f", "Google Drive"],
                                capture_output=True)
        if result.returncode != 0:
            raise Exception("Google Drive not running.")

    elif system == "Windows":
        result = subprocess.run(["tasklist"],
                                capture_output=True,
                                text=True)
        if "GoogleDriveFS" not in result.stdout:
            raise Exception("Google Drive not running.")

    return True
