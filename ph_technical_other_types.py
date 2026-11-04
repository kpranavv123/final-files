"""
PRODUCTHIERARCHY (NON FERT / HAWA) - Technical Validation Tool
Rules 1-4 : field must not be blank
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
INPUT_FILE  = r"C:\Users\SW526XH\Downloads\Go Live-1\ProductH\ProductHierarchy_NonFERT_HAWA.tab"
OUTPUT_FILE = r"C:\Users\SW526XH\Downloads\Go Live-1\ProductH\Validated_Product Hierarchy_Technical_NonFERT_HAWA.xlsx"

# ─────────────────────────────────────────────
#  CONSTANTS & CONFIGURATION
# ─────────────────────────────────────────────

# Material types that belong to the FG ruleset and must NOT be validated here.
# If your input is already a pure non-FERT/HAWA extract this filter simply
# removes nothing. Set APPLY_MATERIALTYPE_FILTER = False to disable it.
APPLY_MATERIALTYPE_FILTER = True
EXCLUDED_MATERIAL_TYPES   = {"FERT", "HAWA"}

# Show only ruleset columns on per-field error sheets (False = show all columns)
SHOW_ONLY_RULESET_COLUMNS = False

# Rules 1-4 : (rule/field name, [candidate column names in the extract])
# The first candidate that exists in the file is used. The second candidate
# is only a fallback in case the extract uses a different spelling.
RULE_FIELDS = [
    ("MATERIALNUMBER",      ["MATERIALNUMBER"]),                       # 1
    ("MATERIALDESCRIPTION", ["MATERIALDESCRIPTION", "PRODUCTDESCRIPTION"]),  # 2
    ("PRODUCTGROUP",        ["PRODUCTGROUP"]),                         # 3
    ("MATLGRPDESC",         ["MATLGRPDESC", "MATLGRPDES"]),            # 4
]

# Filled at runtime: rule/field name -> actual column found in the file
RESOLVED_COLUMNS: dict = {}

NOT_BLANK_FIELDS  = [name for name, _ in RULE_FIELDS]
ALL_SUMMARY_FIELDS = list(NOT_BLANK_FIELDS)

RULES_CONTENT = {f: ["Must not be blank."] for f in NOT_BLANK_FIELDS}

FIELD_REASON_MAP = {f: f"{f}: is blank." for f in NOT_BLANK_FIELDS}

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


def resolve_columns(df: pd.DataFrame):
    """Maps each rule/field name to the actual column present in the file."""
    RESOLVED_COLUMNS.clear()
    for rule_name, candidates in RULE_FIELDS:
        found = next((c for c in candidates if c in df.columns), None)
        if found:
            RESOLVED_COLUMNS[rule_name] = found
            if found != rule_name:
                print(f"    Rule '{rule_name}' → using column '{found}'")
        else:
            print(f"⚠️   None of {candidates} found in input — every row will be flagged for '{rule_name}'")


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
    """Field must not be blank (rules 1-4). A missing column counts as blank."""
    def validate(self, row: pd.Series) -> Optional[str]:
        col = RESOLVED_COLUMNS.get(self.field)
        if col is None or is_blank(row.get(col)):
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

    def validate_dataframe(self, df: pd.DataFrame) -> pd.DataFrame:
        df = df.copy()
        df.columns = df.columns.str.strip().str.upper()

        # ── Keep only NON FERT / HAWA material types ──
        if APPLY_MATERIALTYPE_FILTER:
            if "MATERIALTYPE" in df.columns:
                mtype    = df["MATERIALTYPE"].fillna("").astype(str).str.strip().str.upper()
                excluded = mtype.isin(EXCLUDED_MATERIAL_TYPES)
                print(f"    Rows excluded (FERT / HAWA) : {int(excluded.sum())}")
                df = df[~excluded].copy()
                df["MATERIALTYPE"] = mtype[~excluded]
            else:
                print("⚠️   MATERIALTYPE column not found — material type filter skipped")

        df = df.reset_index(drop=True)
        print(f"    Rows validated (NON FERT / HAWA) : {len(df)}")

        resolve_columns(df)

        error_list        = []
        error_fields_list = []
        error_detail_list = []

        for _, row in df.iterrows():
            # list of (field, reason) pairs
            row_errors: list[tuple[str, str]] = []

            for v in self.validators:
                result = v.validate(row)
                if result:
                    row_errors.append((v.field, result))

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
        """Failed rule names -> actual column names to colour red."""
        return {RESOLVED_COLUMNS.get(f, f) for f in error_fields}

    # ── Sheet 1: PRODUCTHIERARCHY_NON_FG ────────────────

    def _write_main_sheet(self):
        ws         = self.wb.create_sheet("PRODUCTHIERARCHY_NON_FG")
        display_df = self.df.drop(columns=["_ERROR_FIELDS", "_ERROR_DETAIL"])
        headers    = list(display_df.columns)

        ws.append(headers)
        style_header_row(ws, 1, len(headers))
        ws.freeze_panes = "A2"

        for r_idx, (_, row) in enumerate(display_df.iterrows(), start=2):
            failed_rules = self.error_fields.iloc[r_idx - 2]
            has_error    = bool(failed_rules)
            red_columns  = self._highlight_set(failed_rules)

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
        title_cell           = ws.cell(row=1, column=1, value="ProductHierarchy (NON FERT / HAWA) Validation Summary")
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

        # ── Per-field data rows: ALL fields always written ──
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
        title_cell           = ws.cell(row=1, column=1, value="ProductHierarchy (NON FERT / HAWA) – Validation Rules")
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

        all_cols = [c for c in self.df.columns if c not in ("_ERROR_FIELDS", "_ERROR_DETAIL")]

        if SHOW_ONLY_RULESET_COLUMNS:
            ruleset_cols = list(dict.fromkeys(RESOLVED_COLUMNS.values()))
            if "MATERIALTYPE" in all_cols and "MATERIALTYPE" not in ruleset_cols:
                ruleset_cols.append("MATERIALTYPE")
            display_cols = [c for c in ruleset_cols if c in all_cols]
        else:
            display_cols = [c for c in all_cols if c != "ERROR_COLUMNS"]

        # Keep sheet order the same as the rule order
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
            red_col     = RESOLVED_COLUMNS.get(field, field)

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

                if red_col in col_idx_map:
                    target_cell      = ws.cell(row=r_idx, column=col_idx_map[red_col])
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
