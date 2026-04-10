import pandas as pd
import re
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ====================================================
# File paths
# ====================================================
HDA_FILE = r"C:\Users\SW526XH\Downloads\Data Quality Check\tabfiles\HDA(Secondary)_updated.tab"
PART_FILE = r"C:\Users\SW526XH\Downloads\Data Quality Check\Excel_Files\Part.xlsx"
CUSTOMER_FILE = r"C:\Users\SW526XH\Downloads\Data Quality Check\Excel_Files\Customer.xlsx"
OUTPUT_EXCEL = r"C:\Users\SW526XH\Downloads\Data Quality Check\Output_Files\HDA_Secondary_Errors-2.xlsx"

# ====================================================
# Constants
# ====================================================
CHUNK_SIZE = 500_000
EXCEL_MAX_ROWS = 1_048_000

# ====================================================
# Load master data
# ====================================================
part_df = pd.read_excel(PART_FILE, dtype=str)
part_df.columns = part_df.columns.str.strip()
part_set = set(part_df["MATERIALNUMBER"].dropna().str.strip())

customer_df = pd.read_excel(CUSTOMER_FILE, dtype=str)
customer_df.columns = customer_df.columns.str.strip()
customer_set = set(customer_df["CUSTOMER"].dropna().str.strip())

# ====================================================
# Regex
# ====================================================
date_pattern = re.compile(r"^\d{8}$")

# ====================================================
# Error mappings
# ====================================================
ERROR_REASON_MAP = {
    "ERROR_PRODUCT_CODE":      "PRODUCT_CODE: Product code missing in Part master",
    "ERROR_DISTRIBUTOR_CODE":  "DISTRIBUTOR_CODE: Distributor code is blank or missing in Customer master",
    "ERROR_INVOICE_WEEK_START":"INVOICE_WEEK_START: Invoice week start is blank or not in YYYYMMDD format",
}

ERROR_COLUMNS = list(ERROR_REASON_MAP.keys())

SUMMARY_RULES = [
    ("PRODUCT_CODE",      "ERROR_PRODUCT_CODE",      ERROR_REASON_MAP["ERROR_PRODUCT_CODE"]),
    ("DISTRIBUTOR_CODE",  "ERROR_DISTRIBUTOR_CODE",  ERROR_REASON_MAP["ERROR_DISTRIBUTOR_CODE"]),
    ("INVOICE_WEEK_START","ERROR_INVOICE_WEEK_START", ERROR_REASON_MAP["ERROR_INVOICE_WEEK_START"]),
]

# Ruleset descriptions (shown in Rulesets sheet — matches image)
RULESET_DESCRIPTIONS = [
    ("PRODUCT_CODE",      "Must not be blank and must exist in Part master"),
    ("DISTRIBUTOR_CODE",  "Must not be blank and must exist in Customer master"),
    ("INVOICE_WEEK_START","Must not be blank and must be in YYYYMMDD format"),
]

ATTRIBUTE_SHEETS = {
    "ERROR_PRODUCT_CODE":      ("PRODUCT_CODE",       ERROR_REASON_MAP["ERROR_PRODUCT_CODE"]),
    "ERROR_DISTRIBUTOR_CODE":  ("DISTRIBUTOR_CODE",   ERROR_REASON_MAP["ERROR_DISTRIBUTOR_CODE"]),
    "ERROR_INVOICE_WEEK_START":("INVOICE_WEEK_START", ERROR_REASON_MAP["ERROR_INVOICE_WEEK_START"]),
}

attribute_sheet_rows = {
    "PRODUCT_CODE":      {"sheet_no": 1, "row": 0},
    "DISTRIBUTOR_CODE":  {"sheet_no": 1, "row": 0},
    "INVOICE_WEEK_START":{"sheet_no": 1, "row": 0},
}

# ====================================================
# Summary counters
# ====================================================
error_counts = {field: 0 for field, _, _ in SUMMARY_RULES}
total_records = 0
records_with_errors = 0

# ====================================================
# Excel writing setup
# ====================================================
sheet_no    = 1
current_row = 0

with pd.ExcelWriter(OUTPUT_EXCEL, engine="openpyxl") as writer:

    for chunk in pd.read_csv(HDA_FILE, sep="\t", dtype=str, chunksize=CHUNK_SIZE):
        chunk = chunk.apply(lambda x: x.str.strip())
        total_records += len(chunk)

        chunk["ERROR_PRODUCT_CODE"] = chunk["PRODUCT_CODE"].apply(
            lambda x: "Yes" if pd.isna(x) or x == "" or x not in part_set else "")
        chunk["ERROR_DISTRIBUTOR_CODE"] = chunk["DISTRIBUTOR_CODE"].apply(
            lambda x: "Yes" if pd.isna(x) or x == "" or x not in customer_set else "")
        chunk["ERROR_INVOICE_WEEK_START"] = chunk["INVOICE_WEEK_START"].apply(
            lambda x: "Yes" if pd.isna(x) or x == "" or not date_pattern.fullmatch(x) else "")

        chunk["ERROR_COLUMN"] = chunk.apply(
            lambda r: "|".join(
                ERROR_REASON_MAP[c] for c in ERROR_COLUMNS if r[c] == "Yes"
            ), axis=1)

        row_has_error = False
        for field, col, _ in SUMMARY_RULES:
            cnt = (chunk[col] == "Yes").sum()
            error_counts[field] += cnt
            row_has_error = row_has_error | (chunk[col] == "Yes")
        records_with_errors += row_has_error.sum()

        # Consolidated Error_Data sheets
        error_rows = chunk[chunk[ERROR_COLUMNS].eq("Yes").any(axis=1)]
        if not error_rows.empty:
            if current_row + len(error_rows) > EXCEL_MAX_ROWS:
                sheet_no   += 1
                current_row = 0
            error_rows.to_excel(
                writer, sheet_name=f"Error_Data_{sheet_no}",
                startrow=current_row, index=False, header=(current_row == 0))
            current_row += len(error_rows)

        # Attribute-wise error sheets
        for err_col, (base_name, reason_text) in ATTRIBUTE_SHEETS.items():
            attr_rows = chunk[chunk[err_col] == "Yes"].copy()
            if attr_rows.empty:
                continue
            attr_rows["ERROR_COLUMN"] = reason_text
            state = attribute_sheet_rows[base_name]
            if state["row"] + len(attr_rows) > EXCEL_MAX_ROWS:
                state["sheet_no"] += 1
                state["row"]       = 0
            sheet_name = f"{base_name}_{state['sheet_no']}"
            attr_rows.to_excel(
                writer, sheet_name=sheet_name,
                startrow=state["row"], index=False, header=(state["row"] == 0))
            state["row"] += len(attr_rows)

    # ====================================================
    # SUMMARY SHEET
    # ====================================================
    ws_sum = writer.book.create_sheet("Summary")

    title_fill  = PatternFill("solid", fgColor="BDD7EE")
    header_fill = PatternFill("solid", fgColor="D9E1F2")
    total_fill  = PatternFill("solid", fgColor="F2F2F2")
    bold        = Font(bold=True)
    center      = Alignment(horizontal="center")
    thin_side   = Side(style="thin")
    border      = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    ws_sum.merge_cells("A1:G1")
    ws_sum["A1"] = "HDA Secondary Validation Summary"
    ws_sum["A1"].font      = Font(bold=True, size=14)
    ws_sum["A1"].alignment = center
    ws_sum["A1"].fill      = title_fill

    ws_sum.append(["#", "Field Name", "Error Count", "Record Count", "% Health", "% of Error", "Reason"])
    for col in range(1, 8):
        c = ws_sum.cell(row=2, column=col)
        c.font = bold; c.fill = header_fill; c.border = border; c.alignment = center

    row_num = 3
    for idx, (field, _, reason) in enumerate(SUMMARY_RULES, start=1):
        cnt        = error_counts[field]
        pct_error  = round((cnt / total_records) * 100, 2) if total_records else 0
        pct_health = round(100 - pct_error, 2)
        ws_sum.append([idx, field, cnt, total_records, f"{pct_health}%", f"{pct_error}%", reason])
        for col in range(1, 8):
            ws_sum.cell(row=row_num, column=col).border = border
        row_num += 1

    total_errors       = sum(error_counts[f] for f, _, _ in SUMMARY_RULES)
    total_record_count = total_records * len(SUMMARY_RULES)
    total_pct_error    = round((total_errors / total_record_count) * 100, 2) if total_record_count else 0
    total_pct_health   = round(100 - total_pct_error, 2)
    ws_sum.append(["", "TOTAL", total_errors, total_record_count,
                   f"{total_pct_health}%", f"{total_pct_error}%", ""])
    for col in range(1, 8):
        c = ws_sum.cell(row=row_num, column=col)
        c.font = bold; c.fill = total_fill; c.border = border
    row_num += 2

    for label, value in [("Total Records:", total_records),
                         ("Records with Errors:", records_with_errors),
                         ("Records Passing:", total_records - records_with_errors)]:
        ws_sum.cell(row=row_num, column=1).value = label
        ws_sum.cell(row=row_num, column=1).font  = bold
        ws_sum.cell(row=row_num, column=2).value = value
        row_num += 1

    for col_idx, col_cells in enumerate(ws_sum.columns, start=1):
        ws_sum.column_dimensions[get_column_letter(col_idx)].width = (
            max(len(str(c.value)) if c.value else 0 for c in col_cells) + 3)

# ====================================================
# POST-WRITE STYLING  (load workbook back)
# ====================================================
wb = load_workbook(OUTPUT_EXCEL)

pale_yellow_fill = PatternFill("solid", fgColor="FFF2CC")
red_fill         = PatternFill("solid", fgColor="FF0000")
header_blue_fill = PatternFill("solid", fgColor="BDD7EE")
bold             = Font(bold=True)
center           = Alignment(horizontal="center")
thin_side        = Side(style="thin")
border           = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

# Build set of attribute-error sheet base names and their highlight column
#   e.g. "PRODUCT_CODE_1" → highlight column "PRODUCT_CODE"
attribute_base_names = {base: base for base in attribute_sheet_rows}   # base→highlight col

# Collect all attribute error sheet names present in workbook
attr_sheets_in_wb = []
for sname in wb.sheetnames:
    for base in attribute_base_names:
        if sname == f"{base}_1" or (sname.startswith(base + "_") and sname[len(base)+1:].isdigit()):
            attr_sheets_in_wb.append((sname, base))
            break

for sheet_name, base_col in attr_sheets_in_wb:
    ws = wb[sheet_name]

    # Find index of the highlight column from header row
    highlight_col_idx = None
    for cell in ws[1]:
        if cell.value == base_col:
            highlight_col_idx = cell.column
            break

    max_row = ws.max_row
    max_col = ws.max_column

    for row_idx in range(1, max_row + 1):
        for col_idx in range(1, max_col + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            if row_idx == 1:
                # Header row → blue fill + bold + centered
                cell.fill      = header_blue_fill
                cell.font      = Font(bold=True)
                cell.alignment = Alignment(horizontal="center")
                cell.border    = border
            else:
                # Data rows → pale yellow; red on the error column
                if highlight_col_idx is not None and col_idx == highlight_col_idx:
                    cell.fill = red_fill
                else:
                    cell.fill = pale_yellow_fill

wb.save(OUTPUT_EXCEL)
print("✅ Done — Rulesets sheet, error-sheet coloring (pale yellow + red column + blue headers) applied.")

# ====================================================
# ADD RULESETS SHEET  (re-open to keep sheet order tidy)
# ====================================================
wb = load_workbook(OUTPUT_EXCEL)

title_fill  = PatternFill("solid", fgColor="BDD7EE")
header_fill = PatternFill("solid", fgColor="D9E1F2")
green_fill  = PatternFill("solid", fgColor="E2EFDA")
bold        = Font(bold=True)
center      = Alignment(horizontal="center")
thin_side   = Side(style="thin")
border      = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

wsr = wb.create_sheet("Rulesets")

wsr.merge_cells("A1:C1")
wsr["A1"] = "HAD(Secondary) – Validation Rules"
wsr["A1"].font      = Font(bold=True, size=14)
wsr["A1"].fill      = title_fill
wsr["A1"].alignment = center

wsr.append(["#", "Field", "Rule Description"])
for col in range(1, 4):
    c = wsr.cell(row=2, column=col)
    c.font = bold; c.fill = header_fill; c.border = border; c.alignment = center

for idx, (field, rule_desc) in enumerate(RULESET_DESCRIPTIONS, start=1):
    wsr.append([idx, field, rule_desc])
    wsr.cell(row=idx + 2, column=2).fill = green_fill
    for col in range(1, 4):
        wsr.cell(row=idx + 2, column=col).border = border

for col_idx, col_cells in enumerate(wsr.columns, start=1):
    wsr.column_dimensions[get_column_letter(col_idx)].width = (
        max(len(str(c.value)) if c.value else 0 for c in col_cells) + 3)

wb.save(OUTPUT_EXCEL)
print("✅ Rulesets sheet added successfully.")
