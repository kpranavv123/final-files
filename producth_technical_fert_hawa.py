"""
PRODUCTHIERARCHY (FG) - Technical Validation Tool
Rules 1-25 : field must not be blank (FERT / HAWA extract)
Rule  26   : DUPLICATE_CHECK -> MATERIALNUMBER must not be duplicated
Generates a formatted Excel report.
"""

import math
import os
from collections import defaultdict
from typing import Optional

import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ─────────────────────────────────────────────
#  FILE PATHS  –  update these
# ─────────────────────────────────────────────
INPUT_FILE  = r"C:\Users\SW526XH\Downloads\Go Live-1\ProductH\ProductHierarchy_Without_Z3_Materials.tab"
OUTPUT_FILE = r"C:\Users\SW526XH\Downloads\Go Live-1\ProductH\Validated_Product Hierarchy_Technical.xlsx"

# ─────────────────────────────────────────────
#  CONSTANTS & CONFIGURATION
# ─────────────────────────────────────────────

# Rules 1-25 : all use the same rule -> "Must not be blank"
# (extract is already filtered to FERT & HAWA, so no material type filter is needed)
NOT_BLANK_FIELDS = [
    "MATERIALNUMBER",            # 1
    "MATERIALDESCRIPTION",       # 2
    "PRODUCTGROUP",              # 3
    "MATLGRPDESC",               # 4
    "DIVISION",                  # 5
    "DIVISIONDESCRIPTION",       # 6
    "CATEGORY",                  # 7
    "CATEGORYDESCRIPTION",       # 8
    "PRODUCT",                   # 9
    "PRODUCTDESCRIPTION",        # 10
    "VARIANT",                   # 11
    "VARIANTDESCRIPTION",        # 12
    "BRAND",                     # 13
    "BRANDDESCRIPTION",          # 14
    "SUBBRAND",                  # 15
    "SUBBRANDDESCRIPTION",       # 16
    "BRANDVARIANT",              # 17
    "BRANDVARIANTDESCRIPTION",   # 18
    "PACKSIZE",                  # 19
    "PACKSIZEDESCRIPTION",       # 20
    "MARKETSKU",                 # 21
    "MARKETSKUDESCRIPTION",      # 22
    "SUPPLY_FAMILY",             # 23
    "MATERIALTYPE",              # 24
    "IBPSTATUS",                 # 25
]

# Rule 26 (rule name, NOT a column name)
DUPLICATE_RULE_NAME   = "DUPLICATE_CHECK"
DUPLICATE_KEY_COLUMN  = "MATERIALNUMBER"
DUPLICATE_REASON      = "Duplicate value found in MATERIALNUMBER column"

# Defines ALL fields that appear in the Summary (order matters).
ALL_SUMMARY_FIELDS = NOT_BLANK_FIELDS + [DUPLICATE_RULE_NAME]

RULES_CONTENT = {
    **{f: ["Must not be blank."] for f in NOT_BLANK_FIELDS},
    DUPLICATE_RULE_NAME: [
        "MATERIALNUMBER must not contain duplicate values across the extract.",
    ],
}

# Reason text (matches what the Summary 'Reason' column shows)
FIELD_REASON_MAP = {
    **{f: f"{f}: is blank." for f in NOT_BLANK_FIELDS},
    DUPLICATE_RULE_NAME: DUPLICATE_REASON,
}

# Rule name -> actual column(s) to highlight red
HIGHLIGHT_COLUMNS = {
    **{f: [f] for f in NOT_BLANK_FIELDS},
    DUPLICATE_RULE_NAME: [DUPLICATE_KEY_COLUMN],
}

# ─────────────────────────────────────────────
#  Colours
# ─────────────────────────────────────────────
RED_FILL       = PatternFill("solid", start_color="FF0000", end_color="FF0000")
ROW_ERROR_FILL = PatternFill("solid", start_color="FFF2CC", end_color="FFF2CC")
HDR_FILL       = PatternFill("solid", start_color="D9E1F2", end_color="D9E1F2")
RULE_FILL      = PatternFill("solid", start_color="E2EFDA", end_color="E2EFDA")
TITLE_FILL     = PatternFill("solid", start_color="BDD7EE", end_color="BDD7EE")
WHITE_FILL     = PatternFill("solid", start_color="FFFFFF", end_color="FFFFFF")
TOTAL_FILL     = PatternFill("solid", start_color="F2F2F2", end_color="F2F2F2")
STATS_FILL     = PatternFill("solid", start_color="EDEDED", end_color="EDEDED")

HDR_FONT    = Font(bold=True, name="Arial")
BODY_FONT   = Font(name="Arial", size=10)
ERR_FONT    = Font(name="Arial", size=10, bold=True, color="FFFFFF")
THIN_BORDER = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"),  bottom=Side(style="thin"),
)


# ─────────────────────────────────────────────
#  HELPER UTILITIES
# ─────────────────────────────────────────────

def is_blank(value) -> bool:
    if value is None:
        return True
    if isinstance(value, float):
        return math.isnan(value)
    return str(value).strip() == ""


def style_header_row(ws, row: int, num_cols: int):
    for col in range(1, num_cols + 1):
        cell           = ws.cell(row=row, column=col)
        cell.fill      = HDR_FILL
        cell.font      = HDR_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border    = THIN_BORDER


def auto_width(ws, min_w=10, max_w=60):
    for col in ws.columns:
        length = max((len(str(c.value)) if c.value else 0) for c in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max(length + 3, min_w), max_w)


# ─────────────────────────────────────────────
#  VALIDATION ENGINE
# ─────────────────────────────────────────────

class FieldValidator:
    def __init__(self, field: str, rule_description: str):
        self.field            = field
        self.rule_description = rule_description

    def validate(self, row: pd.Series) -> Optional[str]:
        raise NotImplementedError


class NotBlankValidator(FieldValidator):
    """Field must not be blank (rules 1-25)."""
    def validate(self, row: pd.Series) -> Optional[str]:
        if is_blank(row.get(self.field)):
            return FIELD_REASON_MAP[self.field]
        return None


# ─────────────────────────────────────────────
#  VALIDATION ORCHESTRATOR
# ─────────────────────────────────────────────

class ProductHierarchyValidator:

    def __init__(self):
        self.validators: list[FieldValidator] = [
            NotBlankValidator(f, "Field must not be blank") for f in NOT_BLANK_FIELDS
        ]
        # Row indices that hold a duplicated MATERIALNUMBER (filled by _precompute_duplicates)
        self._duplicate_indices: set = set()

    def _precompute_duplicates(self, df: pd.DataFrame):
        """
        Flags EVERY occurrence of a repeated MATERIALNUMBER (keep=False).
        Blank MATERIALNUMBERs are ignored here (they are caught by the blank rule).
        """
        self._duplicate_indices = set()

        if DUPLICATE_KEY_COLUMN not in df.columns:
            return

        key       = df[DUPLICATE_KEY_COLUMN].fillna("").astype(str).str.strip()
        non_blank = key != ""
        dup_mask  = key.duplicated(keep=False) & non_blank
        self._duplicate_indices = set(df.index[dup_mask])

    def validate_dataframe(self, df: pd.DataFrame) -> pd.DataFrame:
        df = df.copy()
        df.columns = df.columns.str.strip().str.upper()

        # Warn about any expected column that is missing from the file
        missing = [c for c in NOT_BLANK_FIELDS if c not in df.columns]
        if missing:
            print(f"⚠️   Columns missing in input (all rows will be flagged blank): {missing}")

        # Tidy MATERIALTYPE only for display consistency
        if "MATERIALTYPE" in df.columns:
            df["MATERIALTYPE"] = df["MATERIALTYPE"].fillna("").str.strip().str.upper()

        self._precompute_duplicates(df)
        print(f"    Duplicate MATERIALNUMBER rows: {len(self._duplicate_indices)}")

        error_list        = []
        error_fields_list = []
        error_detail_list = []

        for idx, row in df.iterrows():
            # list of (field, reason) so no string parsing is needed later
            row_errors: list[tuple[str, str]] = []

            for v in self.validators:
                result = v.validate(row)
                if result:
                    row_errors.append((v.field, result))

            if idx in self._duplicate_indices:
                row_errors.append((DUPLICATE_RULE_NAME, DUPLICATE_REASON))

            error_list.append(" | ".join(r for _, r in row_errors) if row_errors else "")

            field_msgs: dict = {}
            for field, reason in row_errors:
                field_msgs.setdefault(field, []).append(reason)

            error_fields_list.append(set(field_msgs.keys()))
            error_detail_list.append({f: " | ".join(m) for f, m in field_msgs.items()})

        df["ERROR_COLUMNS"] = error_list
        df["_ERROR_FIELDS"] = error_fields_list
        df["_ERROR_DETAIL"] = error_detail_list

        return df


# ─────────────────────────────────────────────
#  EXCEL REPORT BUILDER
# ─────────────────────────────────────────────

class ExcelReportBuilder:

    def __init__(self, df_validated: pd.DataFrame, output_path: str):
        self.df           = df_validated
        self.output       = output_path
        self.wb           = openpyxl.Workbook()
        self.wb.remove(self.wb.active)

        self.error_df     = df_validated[df_validated["ERROR_COLUMNS"] != ""].copy()
        self.error_fields = df_validated["_ERROR_FIELDS"]
        self.error_detail = df_validated["_ERROR_DETAIL"]

    def build(self):
        self._write_main_sheet()
        self._write_summary_sheet()
        self._write_ruleset_sheet()
        self._write_per_field_error_sheets()
        self.wb.save(self.output)
        print(f"\n✅  Report saved → {self.output}")

    @staticmethod
    def _highlight_set(error_fields: set) -> set:
        """Converts failed rule names into the actual column names to colour red."""
        cols = set()
        for f in error_fields:
            cols.update(HIGHLIGHT_COLUMNS.get(f, [f]))
        return cols

    # ── Sheet 1: PRODUCTHIERARCHY_FG ────────────────────

    def _write_main_sheet(self):
        ws         = self.wb.create_sheet("PRODUCTHIERARCHY_FG")
        display_df = self.df.drop(columns=["_ERROR_FIELDS", "_ERROR_DETAIL"])
        headers    = list(display_df.columns)

        ws.append(headers)
        style_header_row(ws, 1, len(headers))
        ws.freeze_panes = "A2"

        for r_idx, (_, row) in enumerate(display_df.iterrows(), start=2):
            failed_rules   = self.error_fields.iloc[r_idx - 2]
            has_error      = bool(failed_rules)
            red_columns    = self._highlight_set(failed_rules)

            for c_idx, col in enumerate(headers, start=1):
                value = row[col]
                if isinstance(value, float) and math.isnan(value):
                    value = None
                cell           = ws.cell(row=r_idx, column=c_idx, value=value)
                cell.font      = BODY_FONT
                cell.border    = THIN_BORDER
                cell.alignment = Alignment(horizontal="center", vertical="center")

                if col in red_columns:
                    cell.fill = RED_FILL
                    cell.font = ERR_FONT
                elif has_error:
                    cell.fill = ROW_ERROR_FILL
                else:
                    cell.fill = WHITE_FILL

        auto_width(ws)
        ws.row_dimensions[1].height = 30

    # ── Sheet 2: Summary ─────────────────────────────────

    def _write_summary_sheet(self):
        ws = self.wb.create_sheet("Summary")

        total_rows          = len(self.df)
        records_with_errors = len(self.error_df)
        records_passing     = total_rows - records_with_errors

        field_counts: dict[str, int] = defaultdict(int)
        for fields_set in self.error_fields:
            for f in fields_set:
                field_counts[f] += 1

        # ── Title ──
        ws.merge_cells("A1:G1")
        title_cell           = ws.cell(row=1, column=1, value="ProductHierarchy FG Validation Summary")
        title_cell.font      = Font(name="Arial", bold=True, size=14)
        title_cell.fill      = TITLE_FILL
        title_cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 24

        # ── Column headers ──
        headers = ["#", "Field Name", "Error Count", "Record Count", "% Health", "% of Error", "Reason"]
        for c_idx, h in enumerate(headers, start=1):
            cell           = ws.cell(row=2, column=c_idx, value=h)
            cell.fill      = TITLE_FILL
            cell.font      = Font(name="Arial", bold=True)
            cell.border    = THIN_BORDER
            cell.alignment = Alignment(horizontal="center", vertical="center")

        # ── Per-field data rows: ALL rules always written ──
        row_num = 3
        for field_num, col_name in enumerate(ALL_SUMMARY_FIELDS, start=1):
            count      = field_counts.get(col_name, 0)
            pct_error  = round((count / total_rows) * 100, 2) if total_rows else 0
            pct_health = round(100 - pct_error, 2)
            reason     = FIELD_REASON_MAP[col_name] if count > 0 else ""

            ws.cell(row=row_num, column=1, value=field_num)
            ws.cell(row=row_num, column=2, value=col_name)
            ws.cell(row=row_num, column=3, value=count)
            ws.cell(row=row_num, column=4, value=total_rows)
            ws.cell(row=row_num, column=5, value=f"{pct_health}%")
            ws.cell(row=row_num, column=6, value=f"{pct_error}%")
            ws.cell(row=row_num, column=7, value=reason)

            for c in range(1, 8):
                cell        = ws.cell(row=row_num, column=c)
                cell.font   = BODY_FONT
                cell.border = THIN_BORDER
                cell.alignment = (
                    Alignment(horizontal="left", vertical="center", wrap_text=True)
                    if c == 7
                    else Alignment(horizontal="center", vertical="center")
                )

            row_num += 1

        # ── TOTAL row ──
        total_errors       = sum(field_counts.values())
        total_record_count = total_rows * len(ALL_SUMMARY_FIELDS)
        total_pct_error    = round((total_errors / total_record_count) * 100, 2) if total_record_count else 0
        total_pct_health   = round(100 - total_pct_error, 2)

        for c_idx, val in enumerate(
            ["", "TOTAL", total_errors, total_record_count,
             f"{total_pct_health}%", f"{total_pct_error}%", ""],
            start=1,
        ):
            cell           = ws.cell(row=row_num, column=c_idx, value=val)
            cell.font      = Font(name="Arial", bold=True)
            cell.fill      = TOTAL_FILL
            cell.border    = THIN_BORDER
            cell.alignment = Alignment(horizontal="center", vertical="center")

        row_num += 2

        # ── Stats block ──
        for label, value in [
            ("Total Records:",       total_rows),
            ("Records with Errors:", records_with_errors),
            ("Records Passing:",     records_passing),
        ]:
            ws.merge_cells(start_row=row_num, start_column=1, end_row=row_num, end_column=2)
            label_cell           = ws.cell(row=row_num, column=1, value=label)
            label_cell.font      = Font(name="Arial", bold=True, size=10)
            label_cell.fill      = STATS_FILL
            label_cell.border    = THIN_BORDER
            label_cell.alignment = Alignment(horizontal="left", vertical="center")

            value_cell           = ws.cell(row=row_num, column=3, value=value)
            value_cell.font      = BODY_FONT
            value_cell.border    = THIN_BORDER
            value_cell.alignment = Alignment(horizontal="center", vertical="center")

            row_num += 1

        auto_width(ws, min_w=8, max_w=70)

    # ── Sheet 3: Rule_Set ────────────────────────────────

    def _write_ruleset_sheet(self):
        ws = self.wb.create_sheet("Rule_Set")

        ws.merge_cells("A1:C1")
        title_cell           = ws.cell(row=1, column=1, value="ProductHierarchy FG – Validation Rules")
        title_cell.font      = Font(name="Arial", bold=True, size=13)
        title_cell.fill      = TITLE_FILL
        title_cell.alignment = Alignment(horizontal="center")
        ws.row_dimensions[1].height = 22

        for c_idx, h in enumerate(["#", "Field", "Rule Description"], start=1):
            cell           = ws.cell(row=3, column=c_idx, value=h)
            cell.fill      = HDR_FILL
            cell.font      = HDR_FONT
            cell.border    = THIN_BORDER
            cell.alignment = Alignment(horizontal="center")

        current_row = 4
        rule_num    = 1

        for field, rules_list in RULES_CONTENT.items():
            num_rules = len(rules_list)

            for r_idx, rule_text in enumerate(rules_list):
                num_cell           = ws.cell(row=current_row, column=1, value=rule_num if r_idx == 0 else "")
                num_cell.font      = Font(name="Arial", size=10, bold=(r_idx == 0))
                num_cell.fill      = RULE_FILL
                num_cell.border    = THIN_BORDER
                num_cell.alignment = Alignment(horizontal="center", vertical="center")

                field_cell           = ws.cell(row=current_row, column=2, value=field if r_idx == 0 else "")
                field_cell.font      = Font(name="Arial", size=10, bold=(r_idx == 0))
                field_cell.fill      = RULE_FILL
                field_cell.border    = THIN_BORDER
                field_cell.alignment = Alignment(horizontal="center", vertical="center")

                desc_cell           = ws.cell(row=current_row, column=3, value=rule_text)
                desc_cell.font      = BODY_FONT
                desc_cell.border    = THIN_BORDER
                desc_cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")

                current_row += 1

            if num_rules > 1:
                s = current_row - num_rules
                e = current_row - 1
                ws.merge_cells(start_row=s, start_column=1, end_row=e, end_column=1)
                ws.merge_cells(start_row=s, start_column=2, end_row=e, end_column=2)

            rule_num += 1

        ws.column_dimensions["A"].width = 6
        ws.column_dimensions["B"].width = 30
        ws.column_dimensions["C"].width = 65

    # ── Per-field error sheets ───────────────────────────

    def _write_per_field_error_sheets(self):
        if self.error_df.empty:
            return

        field_rows: dict[str, list] = defaultdict(list)
        for idx in self.error_df.index:
            for field in self.error_fields.loc[idx]:
                field_rows[field].append(idx)

        display_cols = [c for c in self.df.columns if c not in ("_ERROR_FIELDS", "_ERROR_DETAIL")]

        # Keep sheet order the same as the rule order (not alphabetical)
        ordered_fields = [f for f in ALL_SUMMARY_FIELDS if f in field_rows]

        for field in ordered_fields:
            row_indices = field_rows[field]

            sheet_name = field[:31]
            existing   = [s.title for s in self.wb.worksheets]
            counter    = 1
            base_name  = sheet_name
            while sheet_name in existing:
                sheet_name = f"{base_name[:28]}_{counter}"
                counter   += 1

            ws = self.wb.create_sheet(sheet_name)

            subset = self.df.loc[row_indices, display_cols].copy()
            subset["ERROR_COLUMNS"] = [
                self.error_detail.loc[i].get(field, FIELD_REASON_MAP[field]) for i in subset.index
            ]

            final_cols = list(subset.columns)
            ws.append(final_cols)
            style_header_row(ws, 1, len(final_cols))
            ws.freeze_panes = "A2"

            col_idx_map = {col: i for i, col in enumerate(final_cols, start=1)}
            red_cols    = HIGHLIGHT_COLUMNS.get(field, [field])

            for r_idx, (_, row) in enumerate(subset.iterrows(), start=2):
                for c_idx, col in enumerate(final_cols, start=1):
                    value = row[col]
                    if isinstance(value, float) and math.isnan(value):
                        value = None
                    cell           = ws.cell(row=r_idx, column=c_idx, value=value)
                    cell.font      = BODY_FONT
                    cell.border    = THIN_BORDER
                    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                    cell.fill      = ROW_ERROR_FILL

                # Highlight the column(s) responsible for this rule
                for rc in red_cols:
                    if rc in col_idx_map:
                        target_cell      = ws.cell(row=r_idx, column=col_idx_map[rc])
                        target_cell.fill = RED_FILL
                        target_cell.font = ERR_FONT

            note_row = len(row_indices) + 3
            ws.cell(
                row=note_row, column=1,
                value=f"Total error rows for '{field}': {len(row_indices)}",
            ).font = Font(name="Arial", italic=True, size=9, bold=True)

            auto_width(ws, min_w=10, max_w=60)


# ─────────────────────────────────────────────
#  PIPELINE CONTROLLER
# ─────────────────────────────────────────────

class ValidationPipeline:

    def __init__(self, input_path: str, output_path: str):
        self.input_path  = input_path
        self.output_path = output_path
        self.validator   = ProductHierarchyValidator()

    def run(self):
        print(f"📂  Reading: {self.input_path}")
        df = self._read_input()

        print(f"🔍  Validating {len(df)} rows …")
        df_validated = self.validator.validate_dataframe(df)

        error_count = (df_validated["ERROR_COLUMNS"] != "").sum()
        print(f"⚠️   Errors found in {error_count} row(s)")

        print("📝  Building Excel report …")
        builder = ExcelReportBuilder(df_validated, self.output_path)
        builder.build()

    def _read_input(self) -> pd.DataFrame:
        ext = os.path.splitext(self.input_path)[1].lower()
        if ext in (".xlsx", ".xlsm", ".xls"):
            return pd.read_excel(self.input_path, dtype=str)
        elif ext in (".csv", ".tab"):
            return pd.read_csv(
                self.input_path,
                dtype=str,
                sep="\t",
                encoding="latin-1",
            )
        else:
            raise ValueError(f"Unsupported file type: {ext}")


# ─────────────────────────────────────────────
#  ENTRY POINT
# ─────────────────────────────────────────────

if __name__ == "__main__":
    pipeline = ValidationPipeline(INPUT_FILE, OUTPUT_FILE)
    pipeline.run()
