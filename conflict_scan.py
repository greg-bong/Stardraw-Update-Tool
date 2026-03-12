import pandas as pd

SOURCE_FILE = "/Users/gregoryknight/Downloads/Adyen Standards Documentation - Products.xlsx"

FIELDS = [
    "0. Device ID",
    "PoE",
    "Power",
    "BTU/hr",
    "Weight",
    "Notes"
]

def find_col(df, keyword):
    for c in df.columns:
        if keyword.lower() in str(c).lower():
            return c
    return None

print("\nLoading source workbook...\n")

src_xl = pd.ExcelFile(SOURCE_FILE)

frames = []
for sh in src_xl.sheet_names:
    df = src_xl.parse(sh)
    frames.append(df)

src = pd.concat(frames, ignore_index=True)

model_col = find_col(src, "model")

conflict_total = 0

for field in FIELDS:

    col = find_col(src, field)

    if not col:
        continue

    temp = src[[model_col, col]].dropna()

    temp[col] = temp[col].astype(str).str.strip()
    temp = temp[temp[col] != ""]

    counts = temp.groupby(model_col)[col].nunique()

    conflicts = counts[counts > 1]

    if len(conflicts) > 0:

        print(f"\n===== {field} conflicts =====\n")

        for model in conflicts.index:

            vals = (
                temp[temp[model_col] == model][col]
                .value_counts()
            )

            print(f"\nMODEL: {model}")

            for v, c in vals.items():
                print(f"{v}   [{c} devices]")

            conflict_total += 1

print("\n--------------------------------")
print(f"\nTotal conflicts detected: {conflict_total}")
print("--------------------------------\n")
