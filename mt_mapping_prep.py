import os
from pathlib import Path
import pandas as pd

# ================= CONFIG =================
BASE_DIR = Path(__file__).resolve().parent

RAW_FILE = BASE_DIR / "Raw Data" / "raw_input.xlsx"        # raw HC dump
MAPPING_FILE = BASE_DIR / "Raw Data" / "mapping.xlsx"      # workbook with 2 sheets

RAW_SHEET_NAME = 0                 # first sheet in raw_input
MAPPING_SHEET_NAME = "mapping"     # sheet with BL6 + MT Domain + MT Rollup2 + Generic Dept
EXISTING_SHEET_NAME = "existing"   # sheet with Bank ID + Justification

OUTPUT_FOLDER = BASE_DIR / "Output"
OUTPUT_BASENAME = "HC_output_with_mapping.xlsx"
# =========================================


def get_unique_filename(base_path):
    """
    If base_path exists, append _1, _2, _3 etc.
    """
    base_path = str(base_path)
    if not os.path.exists(base_path):
        return base_path

    name, ext = os.path.splitext(base_path)
    counter = 1
    new_file = f"{name}_{counter}{ext}"

    while os.path.exists(new_file):
        counter += 1
        new_file = f"{name}_{counter}{ext}"

    return new_file


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

    # Core columns
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

    # Extra columns (optional)
    out["Country"] = df_filtered["Country"]
    out["Employment Type"] = df_filtered["Employment Type"]
    out["Global Business Function"] = df_filtered["Global Business Function"]

    return out


def apply_mapping(output_df):
    """
    Use mapping.xlsx to fill MT Domain, Generic Dept (roll up), and Justification.
    """

    # ---- Sheet: mapping ----
    map_df = pd.read_excel(MAPPING_FILE, sheet_name=MAPPING_SHEET_NAME, dtype=str)
    map_df = map_df.fillna("")

    # Build dicts for mapping
    # Business Level 6 Desc -> MT Domain
    domain_map = (
        map_df[["Business Level 6 Desc", "MT Domain"]]
        .drop_duplicates()
        .set_index("Business Level 6 Desc")["MT Domain"]
        .to_dict()
    )

    # MT Rollup Hierarchy 2 Name -> Generic Dept (roll up)
    generic_map = (
        map_df[["MT Rollup Hierarchy 2 Name", "Generic Dept (roll up)"]]
        .drop_duplicates()
        .set_index("MT Rollup Hierarchy 2 Name")["Generic Dept (roll up)"]
        .to_dict()
    )

    output_df["MT Domain"] = output_df["Business Level 6 Desc"].map(domain_map).fillna("")
    output_df["Generic Dept (roll up)"] = (
        output_df["MT Rollup Hierarchy 2 Name"].map(generic_map).fillna("")
    )

    # ---- Sheet: existing ----
    existing_df = pd.read_excel(MAPPING_FILE, sheet_name=EXISTING_SHEET_NAME, dtype=str)
    existing_df = existing_df.fillna("")

    # Bank ID -> Justification
    just_map = (
        existing_df[["Bank ID", "Justification"]]
        .drop_duplicates()
        .set_index("Bank ID")["Justification"]
        .to_dict()
    )

    output_df["Justification"] = output_df["Bank ID"].map(just_map).fillna("")

    return output_df


def main():
    print("Loading raw file...")
    df_raw = load_raw()
    print(f"Raw rows: {len(df_raw)}")

    print("Filtering raw data...")
    df_filtered = filter_raw(df_raw)
    print(f"Rows after filters: {len(df_filtered)}")

    print("Building base output...")
    output = build_base_output(df_filtered)

    print("Applying mapping from mapping.xlsx...")
    output = apply_mapping(output)

    # Ensure output folder
    OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)

    # Unique filename logic
    base_output_path = OUTPUT_FOLDER / OUTPUT_BASENAME
    final_path = get_unique_filename(base_output_path)

    print(f"Saving to: {final_path}")
    output.to_excel(final_path, index=False)
    print("Done.")


if __name__ == "__main__":
    main()
