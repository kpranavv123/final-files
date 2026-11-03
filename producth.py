"""
PRODUCTHIERARCHY (FG) - High-Performance Excel Data Validation Tool
Optimized for 50,000+ rows. Same output structure, much faster generation.
"""

import math
import os
from collections import defaultdict
from typing import Optional, Set, Dict, List

import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows

# ─────────────────────────────────────────────
#  FILE PATHS  –  update these
# ─────────────────────────────────────────────
INPUT_FILE  = r"C:\Users\SW526XH\Downloads\Go Live-1\ProductH\ProductHierarchy_Without_Z3_Materials.tab"
OUTPUT_FILE = r"C:\Users\SW526XH\Downloads\Go Live-1\ProductH\Validated_Product Hierarchy_Technical.xlsx"

# ─────────────────────────────────────────────
#  CONSTANTS & CONFIGURATION
# ─────────────────────────────────────────────
VALID_IBP_STATUSES = {"IBP", "IBP EXP"}

NOT_BLANK_FIELDS = [
    "MATERIALNUMBER", "MATERIALDESCRIPTION", "PRODUCTGROUP", "MATLGRPDESC",
    "DIVISION", "DIVISIONDESCRIPTION", "PRODUCTTYPE",
    "CATEGORY", "CATEGORYDESCRIPTION", "PRODUCT", "PRODUCTDESCRIPTION",
    "VARIANT", "VARIANTDESCRIPTION", "BRAND", "BRANDDESCRIPTION",
    "SUBBRAND", "SUBBRANDDESCRIPTION", "BRANDVARIANT", "BRANDVARIANTDESCRIPTION",
    "PACKSIZE", "PACKSIZEDESCRIPTION", "MARKETSKU", "MARKETSKUDESCRIPTION",
    "SUPPLY_FAMILY",
]

POSITIVE_FIELDS = ["NETWEIGHT", "GROSSWEIGHT", "LENGTH", "WIDTH", "HEIGHT"]

DUPLICATE_SUMMARY_LABEL = "Duplicate value found in MATERIALNUMBER column"

GW_VS_NET_REASON = "GROSSWEIGHT: must be >= NETWEIGHT."

ALL_SUMMARY_FIELDS = (
    NOT_BLANK_FIELDS + POSITIVE_FIELDS + ["MATERIALTYPE", "IBPSTATUS", "DUPLICATE_CHECK"]
)

RULES_CONTENT = {
    **{f: ["Must not be blank."] for f in NOT_BLANK_FIELDS},
    **{f: ["Field must be greater than 0."] for f in POSITIVE_FIELDS},
    "GROSSWEIGHT": [
        "Field must be greater than 0.",
        "GROSSWEIGHT must be >= NETWEIGHT.",
    ],
    "MATERIALTYPE": ["Must not be blank."],
    "IBPSTATUS": [
        "Allowed values: IBP, IBP EXP or blank.",
        "Any other value is treated as an error.",
    ],
    "DUPLICATE_CHECK": [
        "The MATERIALNUMBER column must not contain duplicate values across the extract.",
    ],
}

FIELD_REASON_MAP = {
    **{f: f"{f}: is blank." for f in NOT_BLANK_FIELDS},
    **{f: f"{f}: must be greater than 0." for f in POSITIVE_FIELDS},
    "MATERIALTYPE":    "MATERIALTYPE: is blank.",
    "IBPSTATUS":       "IBPSTATUS: Invalid value - must be 'IBP', 'IBP EXP' or blank.",
    "DUPLICATE_CHECK": DUPLICATE_SUMMARY_LABEL,
}

SUB_ROW_DEFS = {
    "GROSSWEIGHT": [
        ("  ↳ Not greater than 0", "gw_positive",
         FIELD_REASON_MAP["GROSSWEIGHT"], "must be greater than 0"),
        ("  ↳ Less than NETWEIGHT", "gw_vs_net",
         GW_VS_NET_REASON, "must be >= NETWEIGHT"),
    ],
}

# ─────────────────────────────────────────────
#  Colours & Styles (reused – important for speed)
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
CENTER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT_ALIGN   = Alignment(horizontal="left", vertical="center", wrap_text=True)

# ─────────────────────────────────────────────
#  HELPER UTILITIES
# ─────────────────────────────────────────────
def is_blank_series(s: pd.Series) -> pd.Series:
    """Vectorized blank check."""
    return s.isna() | (s.astype(str).str.strip() == "") | (s.astype(str).str.lower() == "nan")

def to_float_series(s: pd.Series) -> pd.Series:
    """Convert series to float, non-numeric → NaN."""
    return pd.to_numeric(s, errors="coerce")

def style_header_row(ws, row: int, num_cols: int):
    for col in range(1, num_cols + 1):
        cell = ws.cell(row=row, column=col)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        cell.alignment = CENTER_ALIGN
        cell.border = THIN_BORDER

def fast_auto_width(ws, df: pd.DataFrame, min_w=10, max_w=45, sample_rows=2000):
    """Fast column width based on header + sample of data (avoids full sheet scan)."""
    sample = df.head(sample_rows)
    for i, col in enumerate(df.columns, start=1):
        header_len = len(str(col))
        try:
            data_len = sample[col].astype(str).str.len().max()
            if pd.isna(data_len):
                data_len = 0
        except Exception:
            data_len = 0
        width = min(max(max(header_len, int(data_len)) + 2, min_w), max_w)
        ws.column_dimensions[get_column_letter(i)].width = width

# ─────────────────────────────────────────────
#  FAST VALIDATION ENGINE (vectorized)
# ─────────────────────────────────────────────
class ProductHierarchyValidator:
    def validate_dataframe(self, df: pd.DataFrame) -> pd.DataFrame:
        df = df.copy()
        df.columns = df.columns.str.strip().str.upper()

        # Ensure key columns exist
        for col in ["MATERIALTYPE", "IBPSTATUS", "MATERIALNUMBER", "NETWEIGHT", "GROSSWEIGHT"]:
            if col not in df.columns:
                df[col] = ""

        df["MATERIALTYPE"] = df["MATERIALTYPE"].fillna("").astype(str).str.strip().str.upper()
        n = len(df)

        # Collect errors per row as list of strings and set of fields
        error_msgs: List[List[str]] = [[] for _ in range(n)]
        error_fields: List[Set[str]] = [set() for _ in range(n)]
        error_detail: List[Dict[str, str]] = [{} for _ in range(n)]

        def add_error(mask: pd.Series, field: str, reason: str):
            idxs = mask[mask].index
            for idx in idxs:
                error_msgs[idx].append(reason)
                error_fields[idx].add(field)
                error_detail[idx][field] = reason  # last one wins; for multi we join later

        # ── 1. Not-blank fields ──
        for field in NOT_BLANK_FIELDS:
            if field not in df.columns:
                continue
            mask = is_blank_series(df[field])
            add_error(mask, field, FIELD_REASON_MAP[field])

        # ── 2. Positive number fields ──
        for field in POSITIVE_FIELDS:
            if field not in df.columns:
                continue
            nums = to_float_series(df[field])
            mask = nums.isna() | (nums <= 0)
            add_error(mask, field, FIELD_REASON_MAP[field])

        # ── 3. MATERIALTYPE blank ──
        mask = is_blank_series(df["MATERIALTYPE"])
        add_error(mask, "MATERIALTYPE", FIELD_REASON_MAP["MATERIALTYPE"])

        # ── 4. IBPSTATUS ──
        raw = df["IBPSTATUS"].fillna("").astype(str).str.strip()
        blank_mask = is_blank_series(df["IBPSTATUS"])
        upper = raw.str.upper()
        invalid = (~blank_mask) & (~upper.isin(VALID_IBP_STATUSES))
        add_error(invalid, "IBPSTATUS", FIELD_REASON_MAP["IBPSTATUS"])

        # ── 5. GROSSWEIGHT >= NETWEIGHT (only when both numeric) ──
        gross = to_float_series(df["GROSSWEIGHT"])
        net   = to_float_series(df["NETWEIGHT"])
        both_ok = gross.notna() & net.notna()
        mask = both_ok & (gross < net)
        # For rows that fail this, append the reason (may already have positive error)
        for idx in mask[mask].index:
            error_msgs[idx].append(GW_VS_NET_REASON)
            error_fields[idx].add("GROSSWEIGHT")
            # Keep both reasons if present
            prev = error_detail[idx].get("GROSSWEIGHT", "")
            error_detail[idx]["GROSSWEIGHT"] = (
                f"{prev} | {GW_VS_NET_REASON}" if prev else GW_VS_NET_REASON
            )

        # ── 6. Duplicates on MATERIALNUMBER (all occurrences flagged) ──
        if "MATERIALNUMBER" in df.columns:
            mat = df["MATERIALNUMBER"].fillna("").astype(str).str.strip()
            # Treat blank as non-duplicate
            non_blank = mat != ""
            dup_mask = non_blank & mat.duplicated(keep=False)
            for idx in dup_mask[dup_mask].index:
                error_msgs[idx].append(DUPLICATE_SUMMARY_LABEL)
                error_fields[idx].add("DUPLICATE_CHECK")
                error_detail[idx]["DUPLICATE_CHECK"] = DUPLICATE_SUMMARY_LABEL

        # Build final columns
        df["ERROR_COLUMNS"] = [" | ".join(msgs) if msgs else "" for msgs in error_msgs]
        df["_ERROR_FIELDS"] = error_fields
        df["_ERROR_DETAIL"] = error_detail

        print(f"    Duplicate MATERIALNUMBER rows: {dup_mask.sum() if 'MATERIALNUMBER' in df.columns else 0}")
        return df

# ─────────────────────────────────────────────
#  EXCEL REPORT BUILDER (optimized)
# ─────────────────────────────────────────────
class ExcelReportBuilder:
    def __init__(self, df_validated: pd.DataFrame, output_path: str):
        self.df = df_validated
        self.output = output_path
        self.wb = openpyxl.Workbook()
        self.wb.remove(self.wb.active)

        self.error_df = df_validated[df_validated["ERROR_COLUMNS"] != ""].copy()
        self.error_fields = df_validated["_ERROR_FIELDS"]
        self.error_detail = df_validated["_ERROR_DETAIL"]

        # Pre-compute which rows have errors (for yellow fill)
        self.has_error = (df_validated["ERROR_COLUMNS"] != "").values

    def build(self):
        print("    Writing main sheet …")
        self._write_main_sheet()
        print("    Writing Summary …")
        self._write_summary_sheet()
        print("    Writing Rule_Set …")
        self._write_ruleset_sheet()
        print("    Writing per-field error sheets …")
        self._write_per_field_error_sheets()
        self.wb.save(self.output)
        print(f"\n✅  Report saved → {self.output}")

    def _get_field_subcounts(self, field_name: str) -> dict:
        defs = SUB_ROW_DEFS.get(field_name, [])
        counts = {key: 0 for _, key, _, _ in defs}
        for detail in self.error_detail:
            reason = detail.get(field_name, "")
            if not reason:
                continue
            for _, key, _, marker in defs:
                if marker in reason:
                    counts[key] += 1
        return counts

    # ── Sheet 1: PRODUCTHIERARCHY_FG (fast bulk write + selective style) ──
    def _write_main_sheet(self):
        ws = self.wb.create_sheet("PRODUCTHIERARCHY_FG")

        display_df = self.df.drop(columns=["_ERROR_FIELDS", "_ERROR_DETAIL"])
        headers = list(display_df.columns)
        n_cols = len(headers)

        # Header
        ws.append(headers)
        style_header_row(ws, 1, n_cols)
        ws.freeze_panes = "A2"
        ws.row_dimensions[1].height = 30

        # Bulk write data rows (very fast)
        # Convert to list of lists once
        data_values = display_df.values.tolist()
        for row_vals in data_values:
            ws.append(row_vals)

        # Apply styling only where needed (this is the key speed optimization)
        # 1. Light yellow on entire error rows
        # 2. Red + white font only on the specific error cells
        for r_idx, (has_err, fields) in enumerate(zip(self.has_error, self.error_fields), start=2):
            if not has_err:
                continue

            # Yellow fill for the whole error row
            for c in range(1, n_cols + 1):
                cell = ws.cell(row=r_idx, column=c)
                cell.fill = ROW_ERROR_FILL
                cell.font = BODY_FONT
                cell.border = THIN_BORDER
                cell.alignment = CENTER_ALIGN

            # Red highlight on the actual error columns
            for field in fields:
                if field == "DUPLICATE_CHECK":
                    # Highlight MATERIALNUMBER for duplicates
                    if "MATERIALNUMBER" in headers:
                        col_idx = headers.index("MATERIALNUMBER") + 1
                        cell = ws.cell(row=r_idx, column=col_idx)
                        cell.fill = RED_FILL
                        cell.font = ERR_FONT
                elif field in headers:
                    col_idx = headers.index(field) + 1
                    cell = ws.cell(row=r_idx, column=col_idx)
                    cell.fill = RED_FILL
                    cell.font = ERR_FONT

                    # Also highlight NETWEIGHT when GW < NET
                    if field == "GROSSWEIGHT" and GW_VS_NET_REASON in str(
                        self.error_detail.iloc[r_idx - 2].get("GROSSWEIGHT", "")
                    ):
                        if "NETWEIGHT" in headers:
                            net_idx = headers.index("NETWEIGHT") + 1
                            net_cell = ws.cell(row=r_idx, column=net_idx)
                            net_cell.fill = RED_FILL
                            net_cell.font = ERR_FONT

        # Light style for clean rows (optional – can comment out for even more speed)
        # We skip full styling of clean rows to stay fast.

        fast_auto_width(ws, display_df)

    # ── Sheet 2: Summary ──
    def _write_summary_sheet(self):
        ws = self.wb.create_sheet("Summary")

        total_rows = len(self.df)
        records_with_errors = len(self.error_df)
        records_passing = total_rows - records_with_errors

        field_counts: Dict[str, int] = defaultdict(int)
        for fields_set in self.error_fields:
            for f in fields_set:
                field_counts[f] += 1

        # Title
        ws.merge_cells("A1:G1")
        title_cell = ws.cell(row=1, column=1, value="ProductHierarchy FG Validation Summary")
        title_cell.font = Font(name="Arial", bold=True, size=14)
        title_cell.fill = TITLE_FILL
        title_cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 24

        # Headers
        headers = ["#", "Field Name", "Error Count", "Record Count", "% Health", "% of Error", "Reason"]
        for c_idx, h in enumerate(headers, start=1):
            cell = ws.cell(row=2, column=c_idx, value=h)
            cell.fill = TITLE_FILL
            cell.font = Font(name="Arial", bold=True)
            cell.border = THIN_BORDER
            cell.alignment = CENTER_ALIGN

        row_num = 3
        for field_num, col_name in enumerate(ALL_SUMMARY_FIELDS, start=1):
            count = field_counts.get(col_name, 0)
            pct_error = round((count / total_rows) * 100, 2) if total_rows else 0
            pct_health = round(100 - pct_error, 2)
            has_sub_rows = col_name in SUB_ROW_DEFS

            reason = ""
            if count > 0 and not has_sub_rows:
                reason = FIELD_REASON_MAP.get(col_name, f"{col_name}: is blank.")

            vals = [field_num, col_name, count, total_rows, f"{pct_health}%", f"{pct_error}%", reason]
            for c, val in enumerate(vals, start=1):
                cell = ws.cell(row=row_num, column=c, value=val)
                cell.font = BODY_FONT
                cell.border = THIN_BORDER
                cell.alignment = LEFT_ALIGN if c == 7 else CENTER_ALIGN
            row_num += 1

            if has_sub_rows and count > 0:
                subcounts = self._get_field_subcounts(col_name)
                for sub_label, sub_key, sub_reason, _ in SUB_ROW_DEFS[col_name]:
                    sub_count = subcounts.get(sub_key, 0)
                    sub_pct_err = round((sub_count / total_rows) * 100, 2) if total_rows else 0
                    sub_pct_health = round(100 - sub_pct_err, 2)
                    vals = ["", sub_label, sub_count, total_rows, f"{sub_pct_health}%", f"{sub_pct_err}%", sub_reason]
                    for c, val in enumerate(vals, start=1):
                        cell = ws.cell(row=row_num, column=c, value=val)
                        cell.font = Font(name="Arial", size=10, italic=True)
                        cell.border = THIN_BORDER
                        if c == 2:
                            cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
                        elif c == 7:
                            cell.alignment = LEFT_ALIGN
                        else:
                            cell.alignment = CENTER_ALIGN
                    row_num += 1

        # TOTAL
        total_errors = sum(field_counts.values())
        total_record_count = total_rows * len(ALL_SUMMARY_FIELDS)
        total_pct_error = round((total_errors / total_record_count) * 100, 2) if total_record_count else 0
        total_pct_health = round(100 - total_pct_error, 2)

        for c_idx, val in enumerate(
            ["", "TOTAL", total_errors, total_record_count, f"{total_pct_health}%", f"{total_pct_error}%", ""],
            start=1,
        ):
            cell = ws.cell(row=row_num, column=c_idx, value=val)
            cell.font = Font(name="Arial", bold=True)
            cell.fill = TOTAL_FILL
            cell.border = THIN_BORDER
            cell.alignment = CENTER_ALIGN
        row_num += 2

        # Stats
        for label, value in [
            ("Total Records:", total_rows),
            ("Records with Errors:", records_with_errors),
            ("Records Passing:", records_passing),
        ]:
            ws.merge_cells(start_row=row_num, start_column=1, end_row=row_num, end_column=2)
            label_cell = ws.cell(row=row_num, column=1, value=label)
            label_cell.font = Font(name="Arial", bold=True, size=10)
            label_cell.fill = STATS_FILL
            label_cell.border = THIN_BORDER
            label_cell.alignment = Alignment(horizontal="left", vertical="center")

            value_cell = ws.cell(row=row_num, column=3, value=value)
            value_cell.font = BODY_FONT
            value_cell.border = THIN_BORDER
            value_cell.alignment = CENTER_ALIGN
            row_num += 1

        # Column widths
        ws.column_dimensions["A"].width = 6
        ws.column_dimensions["B"].width = 28
        ws.column_dimensions["C"].width = 12
        ws.column_dimensions["D"].width = 13
        ws.column_dimensions["E"].width = 11
        ws.column_dimensions["F"].width = 12
        ws.column_dimensions["G"].width = 55

    # ── Sheet 3: Rule_Set ──
    def _write_ruleset_sheet(self):
        ws = self.wb.create_sheet("Rule_Set")
        ws.merge_cells("A1:C1")
        title_cell = ws.cell(row=1, column=1, value="ProductHierarchy FG – Validation Rules")
        title_cell.font = Font(name="Arial", bold=True, size=13)
        title_cell.fill = TITLE_FILL
        title_cell.alignment = Alignment(horizontal="center")
        ws.row_dimensions[1].height = 22

        for c_idx, h in enumerate(["#", "Field", "Rule Description"], start=1):
            cell = ws.cell(row=3, column=c_idx, value=h)
            cell.fill = HDR_FILL
            cell.font = HDR_FONT
            cell.border = THIN_BORDER
            cell.alignment = CENTER_ALIGN

        current_row = 4
        rule_num = 1
        for field, rules_list in RULES_CONTENT.items():
            num_rules = len(rules_list)
            for r_idx, rule_text in enumerate(rules_list):
                num_cell = ws.cell(row=current_row, column=1, value=rule_num if r_idx == 0 else "")
                num_cell.font = Font(name="Arial", size=10, bold=(r_idx == 0))
                num_cell.fill = RULE_FILL
                num_cell.border = THIN_BORDER
                num_cell.alignment = CENTER_ALIGN

                field_cell = ws.cell(row=current_row, column=2, value=field if r_idx == 0 else "")
                field_cell.font = Font(name="Arial", size=10, bold=(r_idx == 0))
                field_cell.fill = RULE_FILL
                field_cell.border = THIN_BORDER
                field_cell.alignment = CENTER_ALIGN

                desc_cell = ws.cell(row=current_row, column=3, value=rule_text)
                desc_cell.font = BODY_FONT
                desc_cell.border = THIN_BORDER
                desc_cell.alignment = CENTER_ALIGN
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

    # ── Per-field error sheets (only errored rows, faster write) ──
    def _write_per_field_error_sheets(self):
        if self.error_df.empty:
            return

        field_rows: Dict[str, List] = defaultdict(list)
        for idx, fields in self.error_fields.items():
            for field in fields:
                field_rows[field].append(idx)

        display_cols = [c for c in self.df.columns if c not in ("_ERROR_FIELDS", "_ERROR_DETAIL")]

        for field, row_indices in sorted(field_rows.items()):
            sheet_name = field[:31]
            existing = {s.title for s in self.wb.worksheets}
            counter = 1
            base_name = sheet_name
            while sheet_name in existing:
                sheet_name = f"{base_name[:28]}_{counter}"
                counter += 1

            ws = self.wb.create_sheet(sheet_name)
            subset = self.df.loc[row_indices, display_cols].copy()

            # Put the specific reason for this field into ERROR_COLUMNS
            default_reason = FIELD_REASON_MAP.get(field, f"{field}: is blank.")
            subset["ERROR_COLUMNS"] = [
                self.error_detail.loc[i].get(field, default_reason) for i in subset.index
            ]

            # Write header + data in bulk
            final_cols = list(subset.columns)
            ws.append(final_cols)
            style_header_row(ws, 1, len(final_cols))
            ws.freeze_panes = "A2"

            for row_vals in subset.values.tolist():
                ws.append(row_vals)

            # Style: yellow row + red on the target column
            col_idx_map = {col: i for i, col in enumerate(final_cols, start=1)}
            target_col = None
            if field == "DUPLICATE_CHECK":
                target_col = col_idx_map.get("MATERIALNUMBER")
            else:
                target_col = col_idx_map.get(field)

            for r_idx in range(2, len(row_indices) + 2):
                for c in range(1, len(final_cols) + 1):
                    cell = ws.cell(row=r_idx, column=c)
                    cell.fill = ROW_ERROR_FILL
                    cell.font = BODY_FONT
                    cell.border = THIN_BORDER
                    cell.alignment = CENTER_ALIGN

                if target_col:
                    cell = ws.cell(row=r_idx, column=target_col)
                    cell.fill = RED_FILL
                    cell.font = ERR_FONT

                # Extra red on NETWEIGHT for GW < NET cases
                if field == "GROSSWEIGHT":
                    reason = str(subset.iloc[r_idx - 2]["ERROR_COLUMNS"])
                    if GW_VS_NET_REASON in reason and "NETWEIGHT" in col_idx_map:
                        net_cell = ws.cell(row=r_idx, column=col_idx_map["NETWEIGHT"])
                        net_cell.fill = RED_FILL
                        net_cell.font = ERR_FONT

            note_row = len(row_indices) + 3
            ws.cell(
                row=note_row, column=1,
                value=f"Total error rows for '{field}': {len(row_indices)}",
            ).font = Font(name="Arial", italic=True, size=9, bold=True)

            fast_auto_width(ws, subset)

# ─────────────────────────────────────────────
#  PIPELINE CONTROLLER
# ─────────────────────────────────────────────
class ValidationPipeline:
    def __init__(self, input_path: str, output_path: str):
        self.input_path = input_path
        self.output_path = output_path
        self.validator = ProductHierarchyValidator()

    def run(self):
        print(f"📂  Reading: {self.input_path}")
        df = self._read_input()
        print(f"🔍  Validating {len(df):,} rows …")
        df_validated = self.validator.validate_dataframe(df)

        error_count = (df_validated["ERROR_COLUMNS"] != "").sum()
        print(f"⚠️   Errors found in {error_count:,} row(s)")

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
                low_memory=False,
            )
        else:
            raise ValueError(f"Unsupported file type: {ext}")

# ─────────────────────────────────────────────
#  ENTRY POINT
# ─────────────────────────────────────────────
if __name__ == "__main__":
    pipeline = ValidationPipeline(INPUT_FILE, OUTPUT_FILE)
    pipeline.run()

