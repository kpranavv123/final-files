import pandas as pd
import re
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
    "ERROR_PRODUCT_CODE": "PRODUCT_CODE:Product code missing in Part master",
    "ERROR_DISTRIBUTOR_CODE": "DISTRIBUTOR_CODE:Distributor code is blank or missing in Customer master",
    "ERROR_INVOICE_WEEK_START": "INVOICE_WEEK_START:Invoice week start is blank or not in YYYYMMDD format"
}

ERROR_COLUMNS = list(ERROR_REASON_MAP.keys())

SUMMARY_RULES = [
    ("PRODUCT_CODE", "ERROR_PRODUCT_CODE", ERROR_REASON_MAP["ERROR_PRODUCT_CODE"]),
    ("DISTRIBUTOR_CODE", "ERROR_DISTRIBUTOR_CODE", ERROR_REASON_MAP["ERROR_DISTRIBUTOR_CODE"]),
    ("INVOICE_WEEK_START", "ERROR_INVOICE_WEEK_START", ERROR_REASON_MAP["ERROR_INVOICE_WEEK_START"]),
]

ATTRIBUTE_SHEETS = {
    "ERROR_PRODUCT_CODE": ("PRODUCT_CODE", ERROR_REASON_MAP["ERROR_PRODUCT_CODE"]),
    "ERROR_DISTRIBUTOR_CODE": ("DISTRIBUTOR_CODE", ERROR_REASON_MAP["ERROR_DISTRIBUTOR_CODE"]),
    "ERROR_INVOICE_WEEK_START": ("INVOICE_WEEK_START", ERROR_REASON_MAP["ERROR_INVOICE_WEEK_START"]),
}

# Track attribute sheet state
attribute_sheet_rows = {
    "PRODUCT_CODE": {"sheet_no": 1, "row": 0},
    "DISTRIBUTOR_CODE": {"sheet_no": 1, "row": 0},
    "INVOICE_WEEK_START": {"sheet_no": 1, "row": 0},
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
sheet_no = 1
current_row = 0

with pd.ExcelWriter(OUTPUT_EXCEL, engine="openpyxl") as writer:

    # ====================================================
    # PROCESS FILE IN CHUNKS
    # ====================================================
    for chunk in pd.read_csv(HDA_FILE, sep="\t", dtype=str, chunksize=CHUNK_SIZE):
        chunk = chunk.apply(lambda x: x.str.strip())
        total_records += len(chunk)

        # -----------------------------
        # Validations
        # -----------------------------
        chunk["ERROR_PRODUCT_CODE"] = chunk["PRODUCT_CODE"].apply(
            lambda x: "Yes" if pd.isna(x) or x == "" or x not in part_set else ""
        )

        chunk["ERROR_DISTRIBUTOR_CODE"] = chunk["DISTRIBUTOR_CODE"].apply(
            lambda x: "Yes" if pd.isna(x) or x == "" or x not in customer_set else ""
        )

        chunk["ERROR_INVOICE_WEEK_START"] = chunk["INVOICE_WEEK_START"].apply(
            lambda x: "Yes" if pd.isna(x) or x == "" or not date_pattern.fullmatch(x) else ""
        )

        # -----------------------------
        # Combined ERROR_COLUMN
        # -----------------------------
        chunk["ERROR_COLUMN"] = chunk.apply(
            lambda r: "|".join(
                ERROR_REASON_MAP[c] for c in ERROR_COLUMNS if r[c] == "Yes"
            ),
            axis=1
        )

        # -----------------------------
        # Summary counters
        # -----------------------------
        row_has_error = False
        for field, col, _ in SUMMARY_RULES:
            cnt = (chunk[col] == "Yes").sum()
            error_counts[field] += cnt
            row_has_error = row_has_error | (chunk[col] == "Yes")

        records_with_errors += row_has_error.sum()

        # -----------------------------
        # Consolidated Error_Data sheets
        # -----------------------------
        error_rows = chunk[chunk[ERROR_COLUMNS].eq("Yes").any(axis=1)]

        if not error_rows.empty:
            if current_row + len(error_rows) > EXCEL_MAX_ROWS:
                sheet_no += 1
                current_row = 0

            error_rows.to_excel(
                writer,
                sheet_name=f"Error_Data_{sheet_no}",
                startrow=current_row,
                index=False,
                header=(current_row == 0)
            )
            current_row += len(error_rows)

        # -----------------------------
        # ATTRIBUTE-WISE ERROR SHEETS
        # -----------------------------
        for err_col, (base_name, reason_text) in ATTRIBUTE_SHEETS.items():

            attr_rows = chunk[chunk[err_col] == "Yes"].copy()
            if attr_rows.empty:
                continue

            attr_rows["ERROR_COLUMN"] = reason_text
            state = attribute_sheet_rows[base_name]

            if state["row"] + len(attr_rows) > EXCEL_MAX_ROWS:
                state["sheet_no"] += 1
                state["row"] = 0

            sheet_name = f"{base_name}_{state['sheet_no']}"

            attr_rows.to_excel(
                writer,
                sheet_name=sheet_name,
                startrow=state["row"],
                index=False,
                header=(state["row"] == 0)
            )

            state["row"] += len(attr_rows)

    # ====================================================
    # SUMMARY SHEET (ONLY ONCE)
    # ====================================================
    ws = writer.book.create_sheet("Summary")

    title_fill = PatternFill("solid", fgColor="BDD7EE")
    header_fill = PatternFill("solid", fgColor="D9E1F2")
    bold = Font(bold=True)
    center = Alignment(horizontal="center")
    border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )

    ws.merge_cells("A1:D1")
    ws["A1"] = "HDA Secondary Validation Summary"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A1"].alignment = center
    ws["A1"].fill = title_fill

    ws.append(["#", "Field Name", "Error Count", "Reason"])

    for col in range(1, 5):
        cell = ws.cell(row=2, column=col)
        cell.font = bold
        cell.fill = header_fill
        cell.border = border
        cell.alignment = center

    row_num = 3
    total_errors = 0

    for idx, (field, _, reason) in enumerate(SUMMARY_RULES, start=1):
        cnt = error_counts[field]
        total_errors += cnt
        ws.append([idx, field, cnt, reason])
        for col in range(1, 5):
            ws.cell(row=row_num, column=col).border = border
        row_num += 1

    ws.append(["", "TOTAL", total_errors, ""])
    for col in range(1, 5):
        ws.cell(row=row_num, column=col).font = bold
        ws.cell(row=row_num, column=col).border = border

    row_num += 2

    ws.append(["Total Records:", total_records])
    ws.append(["Records with Errors:", records_with_errors])
    ws.append(["Records Passing:", total_records - records_with_errors])

    for col_idx, col_cells in enumerate(ws.columns, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = (
            max(len(str(c.value)) if c.value else 0 for c in col_cells) + 3
        )

print("✅ Error data, attribute-wise sheets, and summary created successfully.")
