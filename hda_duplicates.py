import pandas as pd
import sys
import os

# ── Configuration ──────────────────────────────────────────────────────────────
INPUT_FILE       = "input.txt"          # Change to your actual input file path
OUTPUT_DUPES     = "duplicates.txt"     # Output file for duplicate rows
OUTPUT_UNIQUE    = "no_duplicates.txt"  # Output file for unique rows
DUPLICATE_KEY    = ["ITEM", "U_DEPOT", "LOC", "STARTDATE"]
# ───────────────────────────────────────────────────────────────────────────────


def main(input_file=INPUT_FILE, output_dupes=OUTPUT_DUPES, output_unique=OUTPUT_UNIQUE):
    # ── Read input tab-delimited file ──────────────────────────────────────────
    if not os.path.exists(input_file):
        print(f"ERROR: Input file '{input_file}' not found.")
        sys.exit(1)

    df = pd.read_csv(
        input_file,
        sep="\t",
        dtype=str,          # read everything as string to preserve leading zeros etc.
        keep_default_na=False
    )

    print(f"Total rows read       : {len(df)}")

    # ── Validate required columns exist ───────────────────────────────────────
    missing = [col for col in DUPLICATE_KEY if col not in df.columns]
    if missing:
        print(f"ERROR: Missing required columns: {missing}")
        print(f"Columns found in file: {list(df.columns)}")
        sys.exit(1)

    # ── Identify duplicates ────────────────────────────────────────────────────
    # keep=False  →  marks ALL rows of a duplicate group as True
    is_dup = df.duplicated(subset=DUPLICATE_KEY, keep=False)

    df_dupes  = df[is_dup].copy()
    df_unique = df[~is_dup].copy()

    print(f"Duplicate rows        : {len(df_dupes)}")
    print(f"Unique rows           : {len(df_unique)}")

    # ── Write outputs ──────────────────────────────────────────────────────────
    df_dupes.to_csv(output_dupes,  sep="\t", index=False)
    df_unique.to_csv(output_unique, sep="\t", index=False)

    print(f"\nDuplicates written to : {output_dupes}")
    print(f"Unique rows written to: {output_unique}")


if __name__ == "__main__":
    # Optional: accept file paths as command-line arguments
    # Usage: python split_duplicates.py [input] [dupes_out] [unique_out]
    args = sys.argv[1:]
    main(
        input_file   = args[0] if len(args) > 0 else INPUT_FILE,
        output_dupes = args[1] if len(args) > 1 else OUTPUT_DUPES,
        output_unique= args[2] if len(args) > 2 else OUTPUT_UNIQUE,
    )
