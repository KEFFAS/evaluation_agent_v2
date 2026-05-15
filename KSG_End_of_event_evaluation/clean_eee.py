import pandas as pd
import os
from openpyxl.styles import Font

# =========================================================
# EEE CLEANING SCRIPT
# Kenya School of Government - Matuga
# =========================================================

# ===== USER INPUT =====
file_name = input("Enter EEE Excel file name: ").strip()

# =========================================================
# LOAD FILE
# =========================================================
# Load WITHOUT headers first
df = pd.read_excel(file_name, header=None)

print("Original shape:", df.shape)

# =========================================================
# REMOVE FIRST ROW
# =========================================================
# Usually contains export artifacts / metadata
df = df.iloc[1:].reset_index(drop=True)

# =========================================================
# SET TRUE HEADER ROW
# =========================================================
df.columns = df.iloc[0]
df = df[1:].reset_index(drop=True)

# =========================================================
# REMOVE EMPTY ROWS
# =========================================================
df = df.dropna(how='all')

# =========================================================
# REMOVE UNNAMED COLUMNS
# =========================================================
df = df.loc[
    :,
    ~df.columns.astype(str).str.contains("unnamed", case=False)
]

# =========================================================
# CLEAN COLUMN NAMES
# IMPORTANT:
# Preserve original question wording
# =========================================================
df.columns = (
    df.columns
    .astype(str)
    .str.strip()
)

# =========================================================
# CLEAN TEXT VALUES
# =========================================================
for col in df.columns:

    if df[col].dtype == "object":

        df[col] = (
            df[col]
            .astype(str)
            .str.strip()
        )

# =========================================================
# CONVERT NUMERIC COLUMNS
# Smart numeric detection
# =========================================================
for col in df.columns:

    converted = pd.to_numeric(df[col], errors='coerce')

    # If majority values are numeric -> convert
    if converted.notna().sum() > len(df) * 0.5:
        df[col] = converted

# =========================================================
# REMOVE DUPLICATES
# =========================================================
#df = df.drop_duplicates()

# =========================================================
# STANDARDIZE ONLY ESSENTIAL IDENTIFIERS
# DO NOT ALTER QUESTION HEADERS
# =========================================================
rename_map = {
    "Program Name": "Program Title",
    "Coordinator": "Coordinator Name"
}

df = df.rename(columns=rename_map)

# =========================================================
# REMOVE NON-ESSENTIAL SYSTEM COLUMNS
# =========================================================
drop_cols = [
    "Average Rating",
    "Status",
    "Course Duration (Days)"
]

existing_drop_cols = [
    col for col in drop_cols
    if col in df.columns
]

df = df.drop(columns=existing_drop_cols)

# =========================================================
# DETECT IMPORTANT SECTION COLUMNS
# Helps downstream scripts
# =========================================================

# ----- Course Objectives -----
objective_cols = [
    col for col in df.columns
    if "objective" in str(col).lower()
]

# ----- Personal Expectations -----
expectation_cols = [
    col for col in df.columns
    if "expectation" in str(col).lower()
]

# ----- Institution Comparison -----
comparison_cols = [
    col for col in df.columns
    if "similar institution" in str(col).lower()
]

# ----- Qualitative Columns -----
qualitative_cols = [
    col for col in df.columns
    if any(keyword in str(col).lower() for keyword in [
        "suggest",
        "comment",
        "interest",
        "area",
        "training programs"
    ])
]

# =========================================================
# PRINT DETECTED STRUCTURE
# =========================================================
print("\n==============================")
print("Detected Key Sections")
print("==============================")

print("\nCourse Objective Columns:")
for col in objective_cols:
    print("-", col)

print("\nExpectation Columns:")
for col in expectation_cols:
    print("-", col)

print("\nInstitution Comparison Columns:")
for col in comparison_cols:
    print("-", col)

print("\nQualitative Columns:")
for col in qualitative_cols:
    print("-", col)

# =========================================================
# SAVE CLEANED FILE
# =========================================================
base_name = os.path.splitext(file_name)[0]

output_file = f"{base_name}_eee_cleaned.xlsx"

# Prevent overwrite
if os.path.exists(output_file):
    output_file = f"{base_name}_eee_cleaned_new.xlsx"

with pd.ExcelWriter(output_file, engine='openpyxl') as writer:

    df.to_excel(
        writer,
        index=False,
        sheet_name='Cleaned Data'
    )

    worksheet = writer.sheets['Cleaned Data']

    # Bold headers
    for cell in worksheet[1]:
        cell.font = Font(bold=True)

print("\n===================================")
print("✅ Cleaned EEE file saved as:")
print(output_file)
print("===================================")