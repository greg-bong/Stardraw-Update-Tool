# ==========================================
# Stardraw Update Tool Engine
# Version: v1.6.0
# Date: 2026-03-10
# Patch: Numeric formatting + exclusion handling
# ==========================================

import os
import re
import shutil
import sys
import time
import zipfile
from datetime import datetime

import pandas as pd

VERSION = "1.6.0"


def get_resource_path(filename):
    """Resolve bundled resources correctly in both source and packaged app modes."""
    if getattr(sys, "_MEIPASS", None):
        return os.path.join(sys._MEIPASS, filename)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), filename)


EXCLUSION_FILE_PATH = get_resource_path("device_id_exclusions.txt")


# ------------------------------------------------
# Load Device ID exclusion list
# ------------------------------------------------

def load_device_id_exclusions():
    """Load Device ID prefixes that should be cleared instead of normalized."""

    exclusions = set()

    if not os.path.exists(EXCLUSION_FILE_PATH):
        print(f"WARNING: exclusion file not found at {EXCLUSION_FILE_PATH}")
        return exclusions, False

    with open(EXCLUSION_FILE_PATH, "r", encoding="utf-8") as f:
        for line in f:

            val = line.strip().upper()

            if not val:
                continue

            if val.startswith("#"):
                continue

            exclusions.add(val)

    return exclusions, True


DEVICE_ID_EXCLUSIONS, DEVICE_ID_EXCLUSION_FILE_FOUND = load_device_id_exclusions()


def get_device_id_exclusion_status():
    """Return a short status summary for the Device ID exclusion file."""
    return {
        "found": DEVICE_ID_EXCLUSION_FILE_FOUND,
        "path": EXCLUSION_FILE_PATH,
        "count": len(DEVICE_ID_EXCLUSIONS),
    }


# ------------------------------------------------
# Device ID exclusion check
# ------------------------------------------------

def device_id_should_clear(value):
    """Return True when the Device ID prefix is present in the exclusion list."""

    if not value:
        return False

    value = str(value).upper()
    prefix = value.split("-")[0]

    return prefix in DEVICE_ID_EXCLUSIONS


def timestamp():
    """Build a filesystem-safe timestamp for temp and archive filenames."""
    return datetime.now().strftime("%Y-%m-%d_%H-%M-%S")


def find_col(df, name):
    """Find a dataframe column by exact name, ignoring case and outer whitespace."""
    for c in df.columns:
        if c.strip().lower() == name.lower():
            return c
    return None


def normalize_device_id(val):
    """Normalize raw Device IDs to PREFIX-00 unless the prefix is excluded."""

    if pd.isna(val):
        return None

    val = str(val).strip().upper()

    if device_id_should_clear(val):
        return None

    m = re.match(r"([A-Z]{2,3})-", val)

    if m:
        prefix = m.group(1)
        return f"{prefix}-00"

    return None


def resolve_values(series):
    """Collapse a model field to one value, or flag it when conflicting values exist."""

    s = series.dropna()

    vals = s.astype(str).str.strip()
    vals = vals[(vals != "") & (vals.str.lower() != "none")]

    unique_vals = sorted(set(vals))

    if len(unique_vals) == 0:
        return None, False

    if len(unique_vals) == 1:
        return unique_vals[0], False

    return None, True


class AttributeConflictError(Exception):
    """Raised when one or more source attributes contain unresolved conflicting values."""

    def __init__(self, conflicts):
        super().__init__("Attribute conflicts detected")
        self.conflicts = conflicts


# ------------------------------------------------
# Main pipeline
# ------------------------------------------------

def run_pipeline(
    source_path,
    destination_path,
    log_callback,
    progress_callback=None,
    conflict_resolutions=None,
    archive_dir=None,
):
    """Read source attributes, detect conflicts, and write updates into the destination workbook."""

    def log(msg):
        """Send pipeline progress messages back to the UI."""
        log_callback(msg)

    def update_progress(percent, stage):
        """Push structured progress updates back to the GUI when supported."""
        if progress_callback:
            progress_callback(percent, stage)

    update_progress(5, "Starting update")
    log(f"Starting Stardraw Attribute Updater v{VERSION}")
    exclusion_status = get_device_id_exclusion_status()
    if exclusion_status["found"]:
        log(
            "Device ID exclusions loaded: "
            f"{exclusion_status['count']} entr{'y' if exclusion_status['count'] == 1 else 'ies'} "
            f"from {exclusion_status['path']}"
        )
    else:
        log(f"WARNING: Device ID exclusion file not found at {exclusion_status['path']}")

    update_progress(15, "Loading source workbook")
    log("Loading source file...")
    try:
        src_xl = pd.ExcelFile(source_path, engine="openpyxl")
    except zipfile.BadZipFile as exc:
        raise Exception(
            "Source file is not a valid .xlsx workbook. "
            "If it is stored in Google Drive, make sure the file is fully synced and opens locally in Excel."
        ) from exc

    col_keywords = [
        "manufacturer", "model number", "0. device id", "poe", "power", "btu", "weight", "notes"
    ]

    def col_filter(col_name):
        """Keep only source columns that may map into Stardraw attributes."""
        return any(k in str(col_name).strip().lower() for k in col_keywords)

    frames = []

    for sh in src_xl.sheet_names:
        try:
            df = src_xl.parse(sh, usecols=col_filter)
            frames.append(df)
        except ValueError:
            pass

    if not frames:
        raise Exception("Could not find required columns in source file.")

    update_progress(30, "Preparing source data")
    src = pd.concat(frames, ignore_index=True)

    man = find_col(src, "manufacturer")
    model = find_col(src, "model number")
    device = find_col(src, "0. device id")
    poe = find_col(src, "poe")
    power = find_col(src, "power")
    btu = find_col(src, "btu")
    weight = find_col(src, "weight")
    notes = find_col(src, "notes")

    if not model:
        raise Exception("Missing 'model number' column")

    if man:
        src[man] = src[man].astype(str).str.replace(" UDP", "", regex=False)

    for c in [power, poe, btu, weight]:

        if c:
            src[c] = pd.to_numeric(
                src[c].astype(str).str.replace(r"[^\d.\-]", "", regex=True),
                errors="coerce"
            )

    if device:
        src[device] = src[device].apply(normalize_device_id)

    src["_MODEL_DISPLAY"] = src[model].astype(str).str.strip()
    src["_MODEL_NORM"] = src[model].astype(str).str.lower().str.strip()

    fields = {
        "0. Device ID": device,
        "PoE": poe,
        "Power (w)": power,
        "BTU/hr": btu,
        "Weight Kg": weight,
        "Notes": notes
    }

    resolved = {}
    conflicts = []

    update_progress(45, "Scanning models for conflicts")
    log("Scanning models for conflicts...")

    conflict_resolutions = conflict_resolutions or {}

    for m, grp in src.groupby("_MODEL_NORM"):

        resolved[m] = {}

        for field_name, col in fields.items():

            if col not in grp.columns:
                continue

            val, conflict = resolve_values(grp[col])

            if conflict:
                display_name = grp["_MODEL_DISPLAY"].iloc[0]
                options = []

                for raw_value in grp[col].dropna().unique():
                    devices = grp[grp[col].astype(str) == str(raw_value)]

                    if device and device in devices.columns:
                        device_ids = devices[device].dropna().astype(str).tolist()
                    else:
                        device_ids = ["Unknown"]

                    options.append(
                        {
                            "value": str(raw_value),
                            "device_ids": device_ids,
                            "detail": f"{str(raw_value)} -> {', '.join(device_ids)}",
                        }
                    )

                resolution_key = (m, field_name)
                if resolution_key in conflict_resolutions:
                    resolved[m][field_name] = conflict_resolutions[resolution_key]
                else:
                    conflicts.append(
                        {
                            "model_norm": m,
                            "display_name": display_name,
                            "field_name": field_name,
                            "title": f"{display_name} | {field_name} conflict",
                            "options": options,
                        }
                    )
            else:
                resolved[m][field_name] = val

    if conflicts:

        log("Conflicts detected:")

        for conflict in conflicts:
            log("")
            log(conflict["title"])

            for option in conflict["options"]:
                log(option["detail"])

        raise AttributeConflictError(conflicts)

    log("No conflicts detected.")

    update_progress(60, "Checking destination file access")
    log("Testing file access to prevent lock errors...")
    try:
        with open(destination_path, "rb+"):
            pass
    except PermissionError:
        log("ERROR: The destination file is currently locked!")
        raise Exception(
            "Another user has the Stardraw Attributes file open. Please ask them to close it and try again."
        )

    update_progress(70, "Loading destination workbook")
    try:
        attr_xl = pd.ExcelFile(destination_path, engine="openpyxl")
    except zipfile.BadZipFile as exc:
        raise Exception(
            "Destination file is not a valid local .xlsx workbook. "
            "If this is a Google Drive file, make sure you selected the real synced Excel file, not a shortcut or placeholder, "
            "and that it is fully available offline."
        ) from exc

    updated_total = 0

    temp_output_name = f".Stardraw_Attributes_TEMP_{timestamp()}.xlsx"
    temp_output_path = os.path.join(os.path.dirname(destination_path), temp_output_name)

    update_progress(82, "Writing updated workbook")
    log("Building updated model attributes...")

    with pd.ExcelWriter(temp_output_path, engine="openpyxl") as writer:

        for sh in attr_xl.sheet_names:

            df = attr_xl.parse(sh)

            if "Model Number" in df.columns:

                df["_MODEL_NORM"] = df["Model Number"].astype(str).str.lower().str.strip()

                # Destination sheets often infer empty text columns as float64.
                # Cast writable attribute columns to object so string resolutions
                # like Device IDs and Notes can be assigned safely.
                for field_name in fields:
                    if field_name in df.columns:
                        df[field_name] = df[field_name].astype(object)

                for m, vals in resolved.items():

                    matches = df[df["_MODEL_NORM"] == m]

                    for idx in matches.index:

                        row_updated = False

                        for field_name, val in vals.items():

                            if field_name in df.columns and pd.notna(val):

                                df.at[idx, field_name] = val
                                row_updated = True

                        if row_updated:
                            updated_total += 1

                df.drop(columns=["_MODEL_NORM"], inplace=True)

            df.to_excel(writer, index=False, sheet_name=sh)

    if not archive_dir:
        archive_dir = os.path.join(os.path.dirname(destination_path), "Archives")

    os.makedirs(archive_dir, exist_ok=True)

    archive_name = f"ARCHIVE_{timestamp()}_{os.path.basename(destination_path)}"
    archive_path = os.path.join(archive_dir, archive_name)

    update_progress(92, "Archiving and replacing destination file")
    log(f"Creating backup archive at {archive_path}")
    shutil.copy2(destination_path, archive_path)

    log("Replacing live destination workbook...")
    os.replace(temp_output_path, destination_path)

    time.sleep(2)

    update_progress(100, "Update complete")
    log(f"Total rows updated: {updated_total}")
    log("Update complete.")
