import pandas as pd
from pathlib import Path

# ================= CONFIG =================
BASE_DIR = Path(__file__).resolve().parent

RAW_FILE = BASE_DIR / "Raw Data" / "raw_input.xlsx"      # raw HC dump
MAPPING_FILE = BASE_DIR / "Raw Data" / "mapping.xlsx"    # workbook with sheets: mapping, existing

OUTPUT_FILE = BASE_DIR / "Output" / "HC_output_with_mapping.xlsx"
RAW_SHEET_NAME = 0          # first sheet of raw file, change if needed
MAPPING_SHEET_NAME = "mapping"
EXISTING_SHEET_NAME = "existing"
# =========================================


def load_raw():
    df = pd.read_excel(RAW_FILE, sheet_name=RAW_SHEET_NAME, dtype=str)
    df = df.fillna("")
    return df


def filter_raw(df):
    # 1) Global Business Function = Tech and Ops
    mask_gbf = df["Global Business Function"] == "Tech and Ops"

    # 2) Exclude MT Rollup Hierarchy 1 Name = "Eder, Noelle Kathleen"
    mask_not_eder = df["MT Rollup Hierarchy 1 Name"] != "Eder, Noelle Kathleen"

    return df[mask_gbf & mask_not_eder].copy()


def build_base_output(df_filtered):
    """
    Create the basic output columns BEFORE mapping.
    """
    out = pd.DataFrame()

    # Rename core columns
    out["Bank ID"] = df_filtered["Employee ID"]              # Column C in raw
    out["Name"] = df_filtered["Employee Name"]               # Column D
    out["Business Level 6 Desc"] = df_filtered["Business Level 6 Desc"]  # Column V
    out["MT Rollup Hierarchy 1 Name"] = df_filtered["MT Rollup Hierarchy 1 Name"]  # Column BB
    out["MT Rollup Hierarchy 2 Name"] = df_filtered["MT Rollup Hierarchy 2 Name"]  # Column BD

    # Placeholders for mapped fields
    out["MT Domain"] = ""
    out["Generic Dept (roll up)"] = ""
    out["Justification"] = ""

    # Date columns (empty)
    out["Start Date"] = ""
    out["End Date"] = ""

    # Optional extra columns (keep if useful, or drop later)
    out["Country"] = df_filtered["Country"]                  # Column G
    out["Employment Type"] = df_filtered["Employment Type"]  # Column H
    out["Global Business Function"] = df_filtered["Global Business Function"]  # Column Q

    return out


def apply_mapping(output_df):
    """
    Use mapping.xlsx to fill MT Domain, Generic Dept (roll up), and Justification.
    """

    # ---- Sheet: mapping ----
    map_df = pd.read_excel(MAPPING_FILE, sheet_name=MAPPING_SHEET_NAME, dtype=str)
    map_df = map_df.fillna("")

    # For MT Domain: Business Level 6 Desc (col E) -> MT Domain (col F)
    # We assume those columns are named exactly like this in the sheet:
    #   "Business Level 6 Desc" and "MT Domain"
    domain_map = (
        map_df[["Business Level 6 Desc", "MT Domain"]]
        .drop_duplicates()
        .set_index("Business Level 6 Desc")["MT Domain"]
        .to_dict()
    )

    # For Generic Dept (roll up): MT Rollup Hierarchy 2 Name (col A) -> Generic Dept (roll up) (col B)
    # Column names assumed: "MT Rollup Hierarchy 2 Name", "Generic Dept (roll up)"
    generic_map = (
        map_df[["MT Rollup Hierarchy 2 Name", "Generic Dept (roll up)"]]
        .drop_duplicates()
        .set_index("MT Rollup Hierarchy 2 Name")["Generic Dept (roll up)"]
        .to_dict()
    )

    # Map into output
    output_df["MT Domain"] = (
        output_df["Business Level 6 Desc"].map(domain_map).fillna("")
    )
    output_df["Generic Dept (roll up)"] = (
        output_df["MT Rollup Hierarchy 2 Name"].map(generic_map).fillna("")
    )

    # ---- Sheet: existing ----
    existing_df = pd.read_excel(MAPPING_FILE, sheet_name=EXISTING_SHEET_NAME, dtype=str)
    existing_df = existing_df.fillna("")

    # Bank ID (col A) -> Justification (col C)
    # Assuming headers: "Bank ID", "Justification"
    just_map = (
        existing_df[["Bank ID", "Justification"]]
        .drop_duplicates()
        .set_index("Bank ID")["Justification"]
        .to_dict()
    )

    output_df["Justification"] = (
        output_df["Bank ID"].map(just_map).fillna("")
    )

    return output_df


def main():
    print("Loading raw file...")
    df_raw = load_raw()

    print("Filtering raw data...")
    df_filtered = filter_raw(df_raw)

    print("Building base output...")
    output = build_base_output(df_filtered)

    print("Applying mapping from mapping.xlsx...")
    output = apply_mapping(output)

    OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)
    output.to_excel(OUTPUT_FILE, index=False)
    print(f"Done! Saved file to: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
