import pandas as pd
import re
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ======================================================
# File paths
# ======================================================
HDA_FILE = r"C:\Users\SW526XH\Downloads\Data Quality Check\tabfiles\HDA_updated_2.tab"
SUMMARY_HDA_FILE = r"C:\Users\SW526XH\Downloads\Data Quality Check\Output_Files\HDA_VALIDATED_2.tab"

PART_FILE = r"C:\Users\SW526XH\Downloads\Data Quality Check\Excel_Files\Part.xlsx"
CUSTOMER_FILE = r"C:\Users\SW526XH\Downloads\Data Quality Check\Excel_Files\Customer.xlsx"
SITE_FILE = r"C:\Users\SW526XH\Downloads\Data Quality Check\Excel_Files\Site_2026-04-09-1058.csv.xlsx"

OUTPUT_EXCEL = r"C:\Users\SW526XH\Downloads\Data Quality Check\Output_Files\Validated_HDA2.xlsx"

# ======================================================
# Load master/reference data
# ======================================================
part_set = set(pd.read_excel(PART_FILE, dtype=str)["MATERIALNUMBER_PLANT"].dropna().str.strip())
customer_set = set(pd.read_excel(CUSTOMER_FILE, dtype=str)["SUPPLYINGPLANT_CUSTOMER"].dropna().str.strip())
site_set = set(pd.read_excel(SITE_FILE, dtype=str)["PLANT"].dropna().str.strip())

# ======================================================
# Validation rules
# ======================================================
rules = [
    ("PLANT", "ERROR_PLANT", "Must not be blank and must be present in Site master."),
    ("BILLING_WEEK_START", "ERROR_BILLING_WEEK_START", "Must not be blank and must be in YYYYMMDD format."),
    ("MATERIAL_PLANT", "ERROR_MATERIAL_PLANT", "Material-Plant combination must exist in Part master."),
    ("PLANT_SOLDTOPARTY", "ERROR_PLANT_SOLDTOPARTY", "Plant-SoldToParty combination must exist in Customer master.")
]

ERROR_MESSAGES = {col: f"{field}: {reason}" for field, col, reason in rules}

ERROR_SHEETS = {
    "ERROR_PLANT": ("PLANT", ERROR_MESSAGES["ERROR_PLANT"]),
    "ERROR_BILLING_WEEK_START": ("BILLING_WEEK_START", ERROR_MESSAGES["ERROR_BILLING_WEEK_START"]),
    "ERROR_MATERIAL_PLANT": ("MATERIAL_PLANT", ERROR_MESSAGES["ERROR_MATERIAL_PLANT"]),
    "ERROR_PLANT_SOLDTOPARTY": ("PLANT_SOLDTOPARTY", ERROR_MESSAGES["ERROR_PLANT_SOLDTOPARTY"]),
}

# ======================================================
# Constants
# ======================================================
CHUNK_SIZE = 500_000
EXCEL_MAX_ROWS = 1_048_576
date_pattern = re.compile(r"^\d{8}$")

# Track rows and sheet numbers per error
sheet_tracker = {
    sheet: {"sheet_no": 1, "current_row": 0}
    for sheet, _ in ERROR_SHEETS.values()
}

# ======================================================
# WRITE ERROR DATA (SAFE, ALL FEATURES)
# ======================================================
with pd.ExcelWriter(OUTPUT_EXCEL, engine="openpyxl") as writer:

    for chunk in pd.read_csv(HDA_FILE, sep="\t", dtype=str, chunksize=CHUNK_SIZE):
        chunk = chunk.apply(lambda x: x.str.strip())

        # Derived columns
        chunk["MATERIAL"] = chunk["MATERIAL_PLANT"].str.split("_", n=1).str[0]
        chunk["SOLDTOPARTY"] = chunk["PLANT_SOLDTOPARTY"].str.split("_", n=1).str[1]

        # Validations
        chunk["ERROR_PLANT"] = chunk["PLANT"].apply(lambda x: "Yes" if pd.isna(x) or x not in site_set else "")
        chunk["ERROR_BILLING_WEEK_START"] = chunk["BILLING_WEEK_START"].apply(
            lambda x: "Yes" if pd.isna(x) or not date_pattern.match(x) else "")
        chunk["ERROR_MATERIAL_PLANT"] = chunk["MATERIAL_PLANT"].apply(
            lambda x: "Yes" if pd.isna(x) or x not in part_set else "")
        chunk["ERROR_PLANT_SOLDTOPARTY"] = chunk["PLANT_SOLDTOPARTY"].apply(
            lambda x: "Yes" if pd.isna(x) or x not in customer_set else "")

        # Write per error sheet (Excel-safe)
        for error_col, (base_sheet, error_msg) in ERROR_SHEETS.items():
            error_rows = chunk[chunk[error_col] == "Yes"].copy()
            if error_rows.empty:
                continue

            error_rows["ERROR_COLUMNS"] = error_msg
            tracker = sheet_tracker[base_sheet]
            start = 0

            while start < len(error_rows):
                sheet_name = base_sheet if tracker["sheet_no"] == 1 else f"{base_sheet}_{tracker['sheet_no']}"
                remaining = EXCEL_MAX_ROWS - tracker["current_row"]
                write_df = error_rows.iloc[start:start + remaining]

                write_df.to_excel(
                    writer,
                    sheet_name=sheet_name,
                    startrow=tracker["current_row"],
                    index=False,
                    header=(tracker["current_row"] == 0)
                )

                start += len(write_df)
                tracker["current_row"] += len(write_df)

                if tracker["current_row"] >= EXCEL_MAX_ROWS:
                    tracker["sheet_no"] += 1
                    tracker["current_row"] = 0

# ======================================================
# SUMMARY CALCULATION (CHUNK SAFE)
# ======================================================
error_counts = {field: 0 for field, _, _ in rules}
total_records = 0
records_with_errors = 0

for chunk in pd.read_csv(SUMMARY_HDA_FILE, sep="\t", dtype=str, chunksize=CHUNK_SIZE):
    total_records += len(chunk)
    row_has_error = False

    for field, col, _ in rules:
        s = (chunk[col] == "Yes")
        error_counts[field] += s.sum()
        row_has_error |= s

    records_with_errors += row_has_error.sum()

records_passing = total_records - records_with_errors

# ======================================================
# CREATE SUMMARY + RULESETS SHEETS
# ======================================================
wb = load_workbook(OUTPUT_EXCEL)

bold = Font(bold=True)
center = Alignment(horizontal="center")
border = Border(*(Side(style="thin") for _ in range(4)))
title_fill = PatternFill("solid", fgColor="BDD7EE")
header_fill = PatternFill("solid", fgColor="D9E1F2")
green_fill = PatternFill("solid", fgColor="E2EFDA")

# ---------- SUMMARY ----------
ws = wb.create_sheet("Summary")
ws.merge_cells("A1:E1")
ws["A1"] = "HDA Validation Summary"
ws["A1"].font = Font(bold=True, size=14)
ws["A1"].fill = title_fill
ws["A1"].alignment = center

ws.append(["#", "Field Name", "Error Count", "Error %", "Reason"])

for col in range(1, 6):
    c = ws.cell(row=2, column=col)
    c.font = bold
    c.fill = header_fill
    c.border = border
    c.alignment = center

row = 3
for idx, (field, _, reason) in enumerate(rules, start=1):
    cnt = error_counts[field]
    pct = round((cnt / total_records) * 100, 2) if total_records else 0
    ws.append([idx, field, cnt, f"{pct}%", reason])
    for col in range(1, 6):
        ws.cell(row=row, column=col).border = border
    row += 1

row += 1
ws.append(["Total Records", total_records])
ws.append(["Records with Errors", records_with_errors])
ws.append(["Records Passing", records_passing])

# ---------- RULESETS ----------
wsr = wb.create_sheet("Rulesets")
wsr.merge_cells("A1:C1")
wsr["A1"] = "HDA – Validation Rules"
wsr["A1"].font = Font(bold=True, size=14)
wsr["A1"].fill = title_fill
wsr["A1"].alignment = center

wsr.append(["#", "Field", "Rule Description"])

for col in range(1, 4):
    c = wsr.cell(row=2, column=col)
    c.font = bold
    c.fill = header_fill
    c.border = border
    c.alignment = center

row = 3
for idx, (field, _, reason) in enumerate(rules, start=1):
    wsr.append([idx, field, reason])
    wsr.cell(row=row, column=2).fill = green_fill
    for col in range(1, 4):
        wsr.cell(row=row, column=col).border = border
    row += 1

# ---------- Autofit ----------
for sheet in wb.sheetnames:
    wsx = wb[sheet]
    for col_idx, col_cells in enumerate(wsx.columns, start=1):
        wsx.column_dimensions[get_column_letter(col_idx)].width = (
            max(len(str(c.value)) if c.value else 0 for c in col_cells) + 3
        )

wb.save(OUTPUT_EXCEL)

print("✅ ALL FEATURES INCLUDED — script completed successfully")