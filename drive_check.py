import os
import platform
import subprocess

def check_drive(folder_path):
    """Confirm the destination folder exists and Google Drive appears to be running."""
    if not os.path.exists(folder_path):
        raise Exception("Destination folder not found.")

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
