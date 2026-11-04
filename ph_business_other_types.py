import math
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ─────────────────────────────────────────────
#  FILE PATHS
# ─────────────────────────────────────────────
INPUT_FILE  = r"C:\Users\SW526XH\Downloads\Go Live-1\ProductH\ProductHierarchy_NonFERT_HAWA.tab"
OUTPUT_FILE = r"C:\Users\SW526XH\Downloads\Go Live-1\ProductH\Validated_ProductHierarchy_Business_NonFERT_HAWA.xlsx"


# ─────────────────────────────────────────────
#  SCOPE: NON FERT / HAWA MATERIALS
#  FERT and HAWA rows are removed before validation. If the input is already a
#  pure non-FERT/HAWA extract, nothing is removed.
#  Set APPLY_MATERIALTYPE_FILTER = False to validate every row as-is.
# ─────────────────────────────────────────────
APPLY_MATERIALTYPE_FILTER = True
EXCLUDED_MATERIAL_TYPES   = {"FERT", "HAWA"}

# If True, blank / non-numeric values FAIL the "> 0" rules.
# Set to False if blanks should be skipped instead.
BLANK_FAILS_POSITIVE_CHECK = True


# ─────────────────────────────────────────────
#  FIELD RULES  (rule name -> source column it validates)
#  Order here controls Summary / Rules / error-sheet sequence
# ─────────────────────────────────────────────
FIELD_RULES = [
    ("NETWEIGHT_VALUE_CHECK", "NETWEIGHT"),
    ("GROSSWEIGHT",           "GROSSWEIGHT"),
    ("LENGTH",                "LENGTH"),
    ("WIDTH",                 "WIDTH"),
    ("HEIGHT",                "HEIGHT"),
]

FIELD_REASONS = {
    "NETWEIGHT_VALUE_CHECK": "NETWEIGHT: must be greater than 0.",
    "GROSSWEIGHT":           "GROSSWEIGHT: must be greater than 0.",
    "LENGTH":                "LENGTH: must be greater than 0.",
    "WIDTH":                 "WIDTH: must be greater than 0.",
    "HEIGHT":                "HEIGHT: must be greater than 0.",
}
GW_VS_NET_REASON = "GROSSWEIGHT: must be >= NETWEIGHT."

# GROSSWEIGHT has two sub-rules (shown as indented sub-rows in Summary)
GW_SUB_ROWS = [
    ("  ↳ Not greater than 0",  "gw_positive", FIELD_REASONS["GROSSWEIGHT"]),
    ("  ↳ Less than NETWEIGHT", "gw_vs_net",   GW_VS_NET_REASON),
]

FIELD_RULE_LINES = {
    "NETWEIGHT_VALUE_CHECK": [
        "NETWEIGHT must be a number greater than 0.",
        "Blank or non-numeric values are flagged as errors.",
    ],
    "GROSSWEIGHT": [
        "GROSSWEIGHT must be a number greater than 0.",
        "GROSSWEIGHT must be >= NETWEIGHT (evaluated when both values are numeric).",
        "For the >= NETWEIGHT failure, both GROSSWEIGHT and NETWEIGHT cells are highlighted in red.",
    ],
    "LENGTH": [
        "LENGTH must be a number greater than 0.",
        "Blank or non-numeric values are flagged as errors.",
    ],
    "WIDTH": [
        "WIDTH must be a number greater than 0.",
        "Blank or non-numeric values are flagged as errors.",
    ],
    "HEIGHT": [
        "HEIGHT must be a number greater than 0.",
        "Blank or non-numeric values are flagged as errors.",
    ],
}


# ─────────────────────────────────────────────
#  STYLING CONSTANTS  (identical across all scripts)
# ─────────────────────────────────────────────
RED_FILL   = PatternFill("solid", fgColor="FF0000")
ROW_FILL   = PatternFill("solid", fgColor="FFF2CC")
HDR_FILL   = PatternFill("solid", fgColor="D9E1F2")
RULE_FILL  = PatternFill("solid", fgColor="E2EFDA")
TITLE_FILL = PatternFill("solid", fgColor="BDD7EE")
TOTAL_FILL = PatternFill("solid", fgColor="F2F2F2")
WHITE_FILL = PatternFill("solid", fgColor="FFFFFF")
STATS_FILL = PatternFill("solid", fgColor="EDEDED")

HDR_FONT  = Font(bold=True, name="Arial")
BODY_FONT = Font(name="Arial", size=10)
ERR_FONT  = Font(name="Arial", size=10, bold=True, color="FFFFFF")

THIN_BORDER = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"),  bottom=Side(style="thin"),
)


# ══════════════════════════════════════════════
#  Rule Engine
# ══════════════════════════════════════════════
class PHBusinessRuleEngine:
    """
    Weight / dimension field rules:
        NETWEIGHT_VALUE_CHECK, GROSSWEIGHT (> 0 and >= NETWEIGHT), LENGTH, WIDTH, HEIGHT.
    Vectorised (no per-row loops) so it stays fast on large extracts.

    Returns:
        field_error_map : { rule_name : { row_idx : reason_string } }
        gw_sub_map      : { 'gw_positive' : {row_idx: reason},
                            'gw_vs_net'   : {row_idx: reason} }
    """

    @staticmethod
    def _to_numeric(series: pd.Series) -> pd.Series:
        """Strings -> float; blank / non-numeric become NaN."""
        return pd.to_numeric(series.astype(str).str.strip(), errors="coerce")

    def run_field_rules(self, df: pd.DataFrame):
        field_error_map: dict = {}
        gw_sub_map = {"gw_positive": {}, "gw_vs_net": {}}

        numeric = {}
        for rule_name, col in FIELD_RULES:
            if col in df.columns:
                numeric[col] = self._to_numeric(df[col])

        for rule_name, col in FIELD_RULES:
            if col not in numeric:
                print(f"⚠️   Column '{col}' not found — rule '{rule_name}' skipped")
                continue

            num = numeric[col]
            if BLANK_FAILS_POSITIVE_CHECK:
                fail_mask = num.isna() | (num <= 0)
            else:
                fail_mask = num.notna() & (num <= 0)

            rule_errors = {idx: FIELD_REASONS[rule_name] for idx in df.index[fail_mask]}

            # ── GROSSWEIGHT sub-rule (b): gross >= net ──
            if rule_name == "GROSSWEIGHT":
                gw_sub_map["gw_positive"] = dict(rule_errors)

                if "NETWEIGHT" in numeric:
                    gross   = numeric["GROSSWEIGHT"]
                    net     = numeric["NETWEIGHT"]
                    vs_mask = gross.notna() & net.notna() & (gross < net)

                    for idx in df.index[vs_mask]:
                        gw_sub_map["gw_vs_net"][idx] = GW_VS_NET_REASON
                        if idx in rule_errors:
                            rule_errors[idx] = rule_errors[idx] + " | " + GW_VS_NET_REASON
                        else:
                            rule_errors[idx] = GW_VS_NET_REASON
                else:
                    print("⚠️   Column 'NETWEIGHT' not found — GROSSWEIGHT >= NETWEIGHT sub-rule skipped")

            if rule_errors:
                field_error_map[rule_name] = rule_errors

        return field_error_map, gw_sub_map


# ══════════════════════════════════════════════
#  Validator
# ══════════════════════════════════════════════
class PHBusinessValidator:

    def __init__(self, filepath: str):
        self.filepath = filepath

        self.df              = pd.DataFrame()
        self.field_error_map = {}   # { rule_name : { row_idx : reason } }
        self.gw_sub_map      = {"gw_positive": {}, "gw_vs_net": {}}

    @staticmethod
    def _load_file(filepath: str) -> pd.DataFrame:
        path = filepath.lower()
        if path.endswith(".csv"):
            df = pd.read_csv(filepath, dtype=str)
        elif path.endswith(".tab") or path.endswith(".tsv"):
            df = pd.read_csv(
                filepath,
                sep="\t",
                dtype=str,
                encoding="latin-1",
                engine="python",
            )
        elif path.endswith(".xlsx") or path.endswith(".xls"):
            df = pd.read_excel(filepath, dtype=str, engine="openpyxl")
        else:
            raise ValueError(f"Unsupported file format: {filepath}")

        df.columns = [c.strip().upper() for c in df.columns]
        return df

    def load(self):
        df = self._load_file(self.filepath)

        # ── Keep only NON FERT / HAWA material types ──
        if APPLY_MATERIALTYPE_FILTER:
            if "MATERIALTYPE" in df.columns:
                mtype    = df["MATERIALTYPE"].fillna("").astype(str).str.strip().str.upper()
                excluded = mtype.isin(EXCLUDED_MATERIAL_TYPES)
                print(f"    Rows excluded (FERT / HAWA)      : {int(excluded.sum())}")
                df = df[~excluded].copy()
            else:
                print("⚠️   MATERIALTYPE column not found — material type filter skipped")

        self.df = df.reset_index(drop=True)
        print(f"    Rows validated (NON FERT / HAWA) : {len(self.df)}")

    def validate(self):
        engine = PHBusinessRuleEngine()
        self.field_error_map, self.gw_sub_map = engine.run_field_rules(self.df)

    # Convenience: all row indices that have at least one error
    def all_error_indices(self) -> set:
        indices = set()
        for rule_errors in self.field_error_map.values():
            indices.update(rule_errors.keys())
        return indices


# ══════════════════════════════════════════════
#  Report Writer
# ══════════════════════════════════════════════
class PHBusinessReportWriter:

    SHEET_SUMMARY = "Summary"
    SHEET_RULES   = "Rules"

    def __init__(self, validator: PHBusinessValidator, output_path: str):
        self.validator   = validator
        self.output_path = output_path

    # ── Helpers ──────────────────────────────
    def _write_header(self, ws, columns):
        for c_idx, col_name in enumerate(columns, start=1):
            cell           = ws.cell(row=1, column=c_idx, value=col_name)
            cell.fill      = HDR_FILL
            cell.font      = HDR_FONT
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border    = THIN_BORDER
        ws.row_dimensions[1].height = 30

    def _auto_width(self, ws, min_w=10, max_w=60):
        for col in ws.columns:
            length = max((len(str(c.value)) if c.value else 0) for c in col)
            ws.column_dimensions[get_column_letter(col[0].column)].width = min(
                max(length + 3, min_w), max_w
            )

    def _safe_sheet_name(self, wb, name: str) -> str:
        name     = name[:31]
        existing = [s.title for s in wb.worksheets]
        base     = name
        counter  = 1
        while name in existing:
            name = f"{base[:28]}_{counter}"
            counter += 1
        return name

    @staticmethod
    def _clean(value):
        """NaN -> None so Excel cells stay empty."""
        if isinstance(value, float) and math.isnan(value):
            return None
        return value

    # ── Summary row writer ────────────────────
    def _write_summary_row(self, ws, row_num, rule_num, label, count, total,
                           reason, sub=False):
        pct_error  = round((count / total) * 100, 2) if total else 0
        pct_health = round(100 - pct_error, 2)

        ws.cell(row=row_num, column=1, value="" if sub else rule_num)
        ws.cell(row=row_num, column=2, value=label)
        ws.cell(row=row_num, column=3, value=count)
        ws.cell(row=row_num, column=4, value=total)
        ws.cell(row=row_num, column=5, value=f"{pct_health}%")
        ws.cell(row=row_num, column=6, value=f"{pct_error}%")
        ws.cell(row=row_num, column=7, value=reason)

        for c in range(1, 8):
            cell        = ws.cell(row=row_num, column=c)
            cell.font   = Font(name="Arial", size=10, italic=True) if sub else BODY_FONT
            cell.fill   = WHITE_FILL
            cell.border = THIN_BORDER
            if c == 2 and sub:
                cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
            else:
                cell.alignment = Alignment(
                    horizontal="left" if c in (2, 7) else "center",
                    vertical="center",
                    wrap_text=(c == 7),
                )

    # ── Summary Sheet ─────────────────────────
    def _write_summary_sheet(self, wb):
        ws         = wb.create_sheet(self.SHEET_SUMMARY)
        v          = self.validator
        total_rows = len(v.df)

        # ── Title ──
        ws.merge_cells("A1:G1")
        tc           = ws.cell(row=1, column=1,
                               value="Product Hierarchy (NON FERT / HAWA) – Business Rules Validation Summary")
        tc.font      = Font(name="Arial", bold=True, size=14)
        tc.fill      = TITLE_FILL
        tc.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[1].height = 26

        # ── Column headers ──
        headers = ["#", "Rule", "Error Count", "Record Count",
                   "% Health", "% of Error", "Reason"]
        for c_idx, h in enumerate(headers, start=1):
            cell           = ws.cell(row=2, column=c_idx, value=h)
            cell.fill      = TITLE_FILL
            cell.font      = Font(name="Arial", bold=True, size=10)
            cell.border    = THIN_BORDER
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.row_dimensions[2].height = 30

        row_num    = 3
        rule_num   = 1
        all_counts = []

        # ── Field-rule rows ──
        for rule_name, col in FIELD_RULES:
            rule_errors = v.field_error_map.get(rule_name, {})
            count       = len(rule_errors)
            all_counts.append(count)

            has_sub = (rule_name == "GROSSWEIGHT")
            # Fields with sub-rows show their reasons on the sub-rows instead
            reason  = FIELD_REASONS[rule_name] if (count > 0 and not has_sub) else ""

            self._write_summary_row(ws, row_num, rule_num, rule_name, count, total_rows, reason)
            row_num  += 1
            rule_num += 1

            # GROSSWEIGHT sub-rows (only when the field has errors)
            if has_sub and count > 0:
                for sub_label, sub_key, sub_reason in GW_SUB_ROWS:
                    sub_count = len(v.gw_sub_map.get(sub_key, {}))
                    self._write_summary_row(ws, row_num, "", sub_label, sub_count,
                                            total_rows, sub_reason, sub=True)
                    row_num += 1

        # ── TOTAL row ──
        total_errors       = sum(all_counts)
        total_record_count = total_rows * len(FIELD_RULES)
        total_pct_error    = round((total_errors / total_record_count) * 100, 2) if total_record_count else 0
        total_pct_health   = round(100 - total_pct_error, 2)

        for c_idx, val in enumerate(
            ["", "TOTAL", total_errors, total_record_count,
             f"{total_pct_health}%", f"{total_pct_error}%", ""],
            start=1,
        ):
            cell           = ws.cell(row=row_num, column=c_idx, value=val)
            cell.font      = Font(name="Arial", bold=True, size=10)
            cell.fill      = TOTAL_FILL
            cell.border    = THIN_BORDER
            cell.alignment = Alignment(horizontal="center", vertical="center")

        row_num += 2

        # ── Stats block ──
        records_with_errors = len(v.all_error_indices())
        records_passing     = total_rows - records_with_errors

        for label, value in [
            ("Total Records:",       total_rows),
            ("Records with Errors:", records_with_errors),
            ("Records Passing:",     records_passing),
        ]:
            ws.merge_cells(start_row=row_num, start_column=1,
                           end_row=row_num, end_column=2)
            lc           = ws.cell(row=row_num, column=1, value=label)
            lc.font      = Font(name="Arial", bold=True, size=10)
            lc.fill      = STATS_FILL
            lc.border    = THIN_BORDER
            lc.alignment = Alignment(horizontal="left", vertical="center")

            vc           = ws.cell(row=row_num, column=3, value=value)
            vc.font      = BODY_FONT
            vc.border    = THIN_BORDER
            vc.alignment = Alignment(horizontal="center", vertical="center")
            row_num += 1

        # ── Column widths ──
        col_widths = [6, 40, 14, 16, 12, 12, 60]
        for c_idx, width in enumerate(col_widths, start=1):
            ws.column_dimensions[get_column_letter(c_idx)].width = width

    # ── Rules Sheet ───────────────────────────
    def _write_rule_block(self, ws, current_row, rule_num, label, rules_list):
        """Writes one rule block (# / label / description lines). Returns next row."""
        num_rules = len(rules_list)

        for r_idx, rule_text in enumerate(rules_list):
            nc           = ws.cell(row=current_row, column=1,
                                   value=rule_num if r_idx == 0 else "")
            nc.fill      = RULE_FILL
            nc.font      = Font(name="Arial", size=10, bold=(r_idx == 0))
            nc.border    = THIN_BORDER
            nc.alignment = Alignment(horizontal="center", vertical="center")

            fc           = ws.cell(row=current_row, column=2,
                                   value=label if r_idx == 0 else "")
            fc.fill      = RULE_FILL
            fc.font      = Font(name="Arial", size=10, bold=(r_idx == 0))
            fc.border    = THIN_BORDER
            fc.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

            dc           = ws.cell(row=current_row, column=3, value=rule_text)
            dc.font      = BODY_FONT
            dc.border    = THIN_BORDER
            dc.alignment = Alignment(wrap_text=True, vertical="center", horizontal="left")
            current_row += 1

        # Vertically merge # and label columns across rule lines
        if num_rules > 1:
            s = current_row - num_rules
            e = current_row - 1
            ws.merge_cells(start_row=s, start_column=1, end_row=e, end_column=1)
            ws.merge_cells(start_row=s, start_column=2, end_row=e, end_column=2)

        return current_row

    def _write_rules_sheet(self, wb):
        ws = wb.create_sheet(self.SHEET_RULES)

        ws.merge_cells("A1:C1")
        tc           = ws.cell(row=1, column=1,
                               value="Product Hierarchy (NON FERT / HAWA) – Business Validation Rules")
        tc.font      = Font(name="Arial", bold=True, size=13)
        tc.fill      = TITLE_FILL
        tc.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 22

        ws.merge_cells("A2:C2")
        sc           = ws.cell(row=2, column=1,
                               value="Weight and dimension fields must be numeric values greater than 0; "
                                     "GROSSWEIGHT must also be >= NETWEIGHT.")
        sc.font      = Font(name="Arial", italic=True, size=10)
        sc.fill      = RULE_FILL
        sc.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        ws.row_dimensions[2].height = 20

        for c_idx, h in enumerate(["#", "Rule", "Rule Description"], start=1):
            cell           = ws.cell(row=4, column=c_idx, value=h)
            cell.fill      = HDR_FILL
            cell.font      = HDR_FONT
            cell.border    = THIN_BORDER
            cell.alignment = Alignment(horizontal="center", vertical="center")

        current_row = 5
        rule_num    = 1

        for rule_name, _col in FIELD_RULES:
            current_row = self._write_rule_block(
                ws, current_row, rule_num, rule_name, FIELD_RULE_LINES[rule_name],
            )
            rule_num += 1

        ws.column_dimensions["A"].width = 6
        ws.column_dimensions["B"].width = 30
        ws.column_dimensions["C"].width = 75

    # ── Error Sheets — one per field rule ─────
    def _write_field_rule_sheets(self, wb):
        """
        One sheet per field rule that has errors. ALL input columns are shown
        + ERROR_REASON. The rule's own column is highlighted red; for GROSSWEIGHT
        rows failing 'gross >= net', NETWEIGHT is highlighted red as well.
        """
        v              = self.validator
        all_input_cols = list(v.df.columns)

        for rule_name, rule_col in FIELD_RULES:
            rule_errors = v.field_error_map.get(rule_name)
            if not rule_errors:
                continue

            row_indices = list(rule_errors.keys())
            subset      = v.df.loc[row_indices, all_input_cols].copy()
            subset["ERROR_REASON"] = subset.index.map(lambda i: rule_errors.get(i, ""))

            ws = wb.create_sheet(self._safe_sheet_name(wb, rule_name))

            all_cols    = list(subset.columns)
            col_idx_map = {col: i for i, col in enumerate(all_cols, start=1)}

            self._write_header(ws, all_cols)

            for r_idx, (orig_idx, row_data) in enumerate(subset.iterrows(), start=2):
                for col, value in zip(all_cols, row_data):
                    cell           = ws.cell(row=r_idx, column=col_idx_map[col],
                                             value=self._clean(value))
                    cell.font      = BODY_FONT
                    cell.border    = THIN_BORDER
                    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                    cell.fill      = ROW_FILL

                red_cols = [rule_col]
                if rule_name == "GROSSWEIGHT" and orig_idx in v.gw_sub_map.get("gw_vs_net", {}):
                    red_cols.append("NETWEIGHT")

                for rc in red_cols:
                    if rc in col_idx_map:
                        tc      = ws.cell(row=r_idx, column=col_idx_map[rc])
                        tc.fill = RED_FILL
                        tc.font = ERR_FONT

            self._auto_width(ws, min_w=10, max_w=60)
            ws.freeze_panes = "A2"

            note_row = len(row_indices) + 3
            ws.cell(
                row=note_row, column=1,
                value=f"Total error rows for '{rule_name}': {len(row_indices)}",
            ).font = Font(name="Arial", italic=True, size=9, bold=True)

    # ── Orchestrate ───────────────────────────
    def write(self):
        v  = self.validator
        wb = Workbook()

        if "Sheet" in wb.sheetnames:
            del wb["Sheet"]

        self._write_summary_sheet(wb)
        self._write_rules_sheet(wb)
        self._write_field_rule_sheets(wb)

        wb.save(self.output_path)

        print(f"\n✅  Output saved  → {self.output_path}")
        print(f"   Total rows            : {len(v.df)}")
        print(f"   Rows with any error   : {len(v.all_error_indices())}")
        print()
        print("   Field-rule violation counts:")
        for rule_name, _col in FIELD_RULES:
            count = len(v.field_error_map.get(rule_name, {}))
            print(f"     {rule_name:30s} : {count} rows")
        print(f"       ↳ GROSSWEIGHT not > 0       : {len(v.gw_sub_map['gw_positive'])}")
        print(f"       ↳ GROSSWEIGHT < NETWEIGHT   : {len(v.gw_sub_map['gw_vs_net'])}")


# ══════════════════════════════════════════════
#  Orchestrator
# ══════════════════════════════════════════════
class PHBusinessProcessor:

    def __init__(self, input_path: str, output_path: str):
        self.validator = PHBusinessValidator(input_path)
        self.writer    = PHBusinessReportWriter(self.validator, output_path)

    def run(self):
        print("📂  Loading file …")
        self.validator.load()
        print(f"    Columns detected : {list(self.validator.df.columns)}")
        print("🔍  Running field rules …")
        self.validator.validate()
        print("📝  Writing report …")
        self.writer.write()


# ══════════════════════════════════════════════
#  ENTRY POINT
# ══════════════════════════════════════════════
if __name__ == "__main__":
    processor = PHBusinessProcessor(
        input_path  = INPUT_FILE,
        output_path = OUTPUT_FILE,
    )
    processor.run()
