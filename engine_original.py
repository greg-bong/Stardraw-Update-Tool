# ==========================================
# Stardraw Update Tool Engine
# Version: v1.5.0
# Date: 2026-03-06
# Patch: Numeric dtype stabilisation
# ==========================================

import pandas as pd
import os
import shutil
import time
import re
from datetime import datetime

# ------------------------------------------------
# Load Device ID exclusion list
# ------------------------------------------------

def load_device_id_exclusions():

    exclusions = set()

    try:
        with open("device_id_exclusions.txt") as f:
            for line in f:
                val = line.strip().upper()

                if not val:
                    continue

                if val.startswith("#"):
                    continue

                exclusions.add(val)

    except FileNotFoundError:
        pass

    return exclusions


DEVICE_ID_EXCLUSIONS = load_device_id_exclusions()

# ------------------------------------------------
# Check if Device ID should be cleared
# ------------------------------------------------

def device_id_should_clear(value):

    if not value:
        return False

    value = str(value).upper()

    prefix = value.split("-")[0]

    return prefix in DEVICE_ID_EXCLUSIONS


VERSION = "1.5.0"


def timestamp():
    return datetime.now().strftime("%Y-%m-%d_%H-%M-%S")


def find_col(df, name):
    for c in df.columns:
        if c.strip().lower() == name.lower():
            return c
    return None


def normalize_device_id(val):

    if pd.isna(val):
        return None

    val = str(val).strip().upper()

    # check exclusion list
    if device_id_should_clear(val):
        return None

m = re.match(r"([A-Z]{2,3})-", val)

if m:
    prefix = m.group(1)
    return f"{prefix}-00"

return None


def resolve_values(series):

vals = series.dropna().astype(str).str.strip()

vals = vals[vals != ""]

unique_vals = sorted(set(vals))

if len(unique_vals) == 0:
    return None, False

if len(unique_vals) == 1:
    return unique_vals[0], False

return None, True


def run_pipeline(source_path, destination_path, log_callback):

def log(msg):
    log_callback(msg)

log(f"Starting Stardraw Attribute Updater v{VERSION}")

log("Loading source file...")
src_xl = pd.ExcelFile(source_path)

frames = []
for sh in src_xl.sheet_names:
    df = src_xl.parse(sh)
    frames.append(df)

src = pd.concat(frames, ignore_index=True)

man = find_col(src, "manufacturer")
model = find_col(src, "model number")
device = find_col(src, "0. device id")
poe = find_col(src, "poe")
power = find_col(src, "power")
btu = find_col(src, "btu")
weight = find_col(src, "weight")
notes = find_col(src, "notes")

if man:
    src[man] = src[man].astype(str).str.replace(" UDP", "", regex=False)

for c in [power, poe, btu, weight]:
    if c:
        src[c] = (
            src[c]
            .astype(str)
            .str.replace(r"[^\d.\-]", "", regex=True)
        )
        src[c] = pd.to_numeric(src[c], errors="coerce")

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

log("Scanning models for conflicts...")

for m, grp in src.groupby("_MODEL_NORM"):

    resolved[m] = {}

    for field_name, col in fields.items():

        if col not in grp.columns:
            continue

        val, conflict = resolve_values(grp[col])

        if conflict:
            conflicts.append((m, field_name, grp[col].dropna().unique()))
        else:
            resolved[m][field_name] = val

if conflicts:

    log("Conflicts detected:")
    for m, field, vals in conflicts:

        grp = src[src["_MODEL_NORM"] == m]

        display_name = grp["_MODEL_DISPLAY"].iloc[0]

        log("")
        log(f"{display_name} | {field} conflict")

        for v in vals:

            devices = grp[grp[field].astype(str) == str(v)]

            device_ids = devices["0. Device ID"].dropna().astype(str).tolist()

            log(f"{v}   → {', '.join(device_ids)}")

    log("Resolve conflicts before running updater.")
    raise Exception("Attribute conflicts detected")

log("No conflicts detected.")

attr_xl = pd.ExcelFile(destination_path)

updated_total = 0

dated_output_name = f"Stardraw_Attributes_FINAL_{timestamp()}.xlsx"
generated_output_path = os.path.join(os.path.dirname(destination_path), dated_output_name)

with pd.ExcelWriter(generated_output_path, engine="openpyxl") as writer:

    for sh in attr_xl.sheet_names:

        df = attr_xl.parse(sh)

        if "Model Number" in df.columns:

            df["_MODEL_NORM"] = df["Model Number"].astype(str).str.lower().str.strip()

            for m, vals in resolved.items():

                matches = df[df["_MODEL_NORM"] == m]

                for idx in matches.index:

                    for field, val in vals.items():

                        if field in df.columns and val is not None:

                            df.at[idx, field] = val

                    updated_total += 1

            df.drop(columns=["_MODEL_NORM"], inplace=True)

        df.to_excel(writer, index=False, sheet_name=sh)

archive_name = f"ARCHIVE_{timestamp()}_{os.path.basename(destination_path)}"
archive_path = os.path.join(os.path.dirname(destination_path), archive_name)

shutil.copy2(destination_path, archive_path)

shutil.copy2(generated_output_path, destination_path)

log("Waiting for Drive sync...")
time.sleep(5)

log(f"Total rows updated: {updated_total}")
log("Update complete.")
