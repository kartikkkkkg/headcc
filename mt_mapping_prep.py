import pandas as pd

# === CONFIG ===
INPUT_FILE = "raw_input.xlsx"        # change to your raw file
OUTPUT_FILE = "mt_mapping_template.xlsx"   # output file name
SHEET_NAME = 0  # 0 = first sheet, or use a name like "Sheet1"

def main():
    # Read the raw Excel
    df = pd.read_excel(INPUT_FILE, sheet_name=SHEET_NAME, dtype=str)

    # Make sure we don't get NaN; easier to work with empty strings
    df = df.fillna("")

    # --- Filters ---

    # 1) Global Business Function = "Tech and Ops"
    mask_gbf = df["Global Business Function"] == "Tech and Ops"

    # 2) MT Rollup Hierarchy 1 Name != "Eder, Noelle Kathleen"
    mask_not_eder = df["MT Rollup Hierarchy 1 Name"] != "Eder, Noelle Kathleen"

    df_filtered = df[mask_gbf & mask_not_eder].copy()

    # --- Build output columns ---

    # Rename / map columns:
    # Column C  -> Employee ID      -> Bank ID
    # Column D  -> Employee Name    -> Name
    # Column V  -> Business Level 6 Desc
    # Column BB -> MT Rollup Hierarchy 1 Name
    # Column BD -> MT Rollup Hierarchy 2 Name

    output = pd.DataFrame()

    # Core requested columns
    output["Bank ID"] = df_filtered["Employee ID"]
    output["Name"] = df_filtered["Employee Name"]
    output["Business Level 6 Desc"] = df_filtered["Business Level 6 Desc"]
    output["MT Rollup Hierarchy 1 Name"] = df_filtered["MT Rollup Hierarchy 1 Name"]
    output["MT Rollup Hierarchy 2 Name"] = df_filtered["MT Rollup Hierarchy 2 Name"]

    # Columns to be filled later via mapping
    output["MT Domain"] = ""
    output["Generic Dept (roll up)"] = ""
    output["Justification"] = ""

    # Empty date columns
    output["Start Date"] = ""
    output["End Date"] = ""

    # Helpful columns for mapping logic later (can be dropped if you don't want them)
    output["Country"] = df_filtered["Country"]            # Column G
    output["Employment Type"] = df_filtered["Employment Type"]  # Column H
    output["Global Business Function"] = df_filtered["Global Business Function"]  # Column Q

    # Save to Excel
    output.to_excel(OUTPUT_FILE, index=False)
    print(f"Done! Saved filtered template to: {OUTPUT_FILE}")

if __name__ == "__main__":
    main()
