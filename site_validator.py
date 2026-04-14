import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ─────────────────────────────────────────────
#  FILE PATHS
# ─────────────────────────────────────────────
SITE_INPUT_FILE  = r"C:\Users\SW526XH\Downloads\Go Live-1\Site\Site_2026-04-09-1058.csv.xlsx"
PART_INPUT_FILE  = r"C:\Users\SW526XH\Downloads\Go Live-1\Part\Part_site_FG 10.04.2026.csv"
OUTPUT_FILE      = r"C:\Users\SW526XH\Downloads\Go Live-1\Site\Validated_Site.xlsx"


# ─────────────────────────────────────────────
#  CONSOLIDATED PL LIST
# ─────────────────────────────────────────────
VALID_PLANTS = [
    "1127", "1100", "1105", "1146", "1156", "1107", "1157", "1158", "1166", "1180",
    "1184", "1186", "1197", "1203", "1204", "1211", "1213", "1214", "1218", "1223",
    "1110", "1225", "1226", "1229", "1233", "1234", "1240", "1113", "1248", "1253",
    "1257", "1258", "1265", "1114", "1145", "1275", "1279", "1416", "1421", "1423",
    "1425", "1426", "1428", "1429", "1430", "1432", "1433", "1436", "1437", "1438",
    "1439", "1440", "1442", "1445", "1445", "1449", "1451", "1452", "1455", "1463",
    "1471", "1473", "1475", "1476", "1477", "1478", "1480", "1481", "1483", "1484",
    "1485", "1487", "1488", "1233", "1491", "1495", "1500", "1501", "1505", "1233",
    "1506", "1507", "1421", "1508", "1509", "1521", "1525", "1563", "1578", "1650",
    "1651", "1652", "1654", "1656", "1657", "1658", "1659", "1661", "1579", "1627",
    "1589", "1623", "1509", "1646", "1647", "1640", "1112", "5011637", "5123296",
    "5123742", "4007430", "1623", "1722", "1724", "1725", "1726", "1731", "1732",
    "1733", "1734", "1738", "1739", "1740", "1742", "1754", "1757", "1758", "1771",
    "1774", "1780", "1784", "1785", "1788", "1511", "1512", "1448", "4011702",
    "5013796", "5018849", "5011407", "5015073", "5011308", "5123531", "1642", "1638",
    "1104", "1104A", "1643", "1642", "1106", "1647", "1643", "1109", "1645",
    "1645", "1111", "1648", "1648", "1649", "1649", "1653", "1653", "1801", "1802",
    "2082", "2091", "2088", "2089", "2083", "2084", "2085", "2090", "2081",
    "2086", "2087", "1554", "1520", "1472", "1571", "1298", "1295", "1196",
    "1292", "1176", "1441", "1522", "1494", "1489", "1518", "1249", "1296",
    "1137", "1208", "1155", "1569", "1235", "1281", "1503", "1482", "1135",
    "1205", "1241", "1499", "1462", "1555", "1559", "1575", "2092", "1558",
    "1601", "1125", "1256", "1568"
]

VALID_COMPANY_CODES = {"1001", "1006", "1009"}

KEEP_COLS = ["PLANT", "NAME", "ADDRESS", "TCPL_PLANTTYPE", "COMPANYCODE"]

# ─────────────────────────────────────────────
#  Colours / Styles
# ─────────────────────────────────────────────
RED_FILL       = PatternFill("solid", start_color="FF0000",  end_color="FF0000")
ROW_FILL       = PatternFill("solid", start_color="FFF2CC",  end_color="FFF2CC")
HDR_FILL       = PatternFill("solid", start_color="D9E1F2",  end_color="D9E1F2")
RULE_FILL      = PatternFill("solid", start_color="E2EFDA",  end_color="E2EFDA")
TITLE_FILL     = PatternFill("solid", start_color="BDD7EE",  end_color="BDD7EE")
TOTAL_FILL     = PatternFill("solid", start_color="F2F2F2",  end_color="F2F2F2")
WHITE_FILL     = PatternFill("solid", start_color="FFFFFF",  end_color="FFFFFF")
STATS_FILL     = PatternFill("solid", start_color="EDEDED",  end_color="EDEDED")
PLANT_SUB_FILL = PatternFill("solid", start_color="DAEEF3",  end_color="DAEEF3")
NO_ERR_FILL    = PatternFill("solid", start_color="E2EFDA",  end_color="E2EFDA")   # green tint for zero-error rows

HDR_FONT    = Font(bold=True, name="Arial")
BODY_FONT   = Font(name="Arial", size=10)
ERR_FONT    = Font(name="Arial", size=10, bold=True, color="FFFFFF")
THIN_BORDER = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"),  bottom=Side(style="thin"),
)

# ── Canonical field order (drives summary rows AND error-sheet tab order) ──
FIELD_ORDER = ["PLANT", "NAME", "ADDRESS", "TCPL_PLANTTYPE", "COMPANYCODE"]


# ══════════════════════════════════════════════
#  Rule Engine
# ══════════════════════════════════════════════
class SiteRuleEngine:

    def __init__(self, valid_plants: list, part_plants: set):
        self.valid_plants = set(str(p).strip() for p in valid_plants)
        self.part_plants  = set(str(p).strip() for p in part_plants)

    @staticmethod
    def _is_blank(value) -> bool:
        return pd.isna(value) or str(value).strip() == ""

    def validate_plant(self, row) -> str:
        val = str(row.get("PLANT", "")).strip()
        if not val or val == "nan":
            return "PLANT: Field is blank — plant code is mandatory"
        if val not in self.valid_plants:
            return f"PLANT: '{val}' is not present in the Consolidated PL list"
        if val not in self.part_plants:
            return f"PLANT: '{val}' has no active Part-Site combination in the Part master table"
        return ""

    def validate_name(self, row) -> str:
        if self._is_blank(row.get("NAME")):
            return "NAME: Field is blank — site name is mandatory"
        return ""

    def validate_address(self, row) -> str:
        if self._is_blank(row.get("ADDRESS")):
            return "ADDRESS: Field is blank — address is mandatory"
        return ""

    def validate_tcpl_planttype(self, row) -> str:
        if self._is_blank(row.get("TCPL_PLANTTYPE")):
            return "TCPL_PLANTTYPE: Field is blank"
        return ""

    def validate_companycode(self, row) -> str:
        val = row.get("COMPANYCODE", None)
        if self._is_blank(val):
            return "COMPANYCODE: Field is blank — company code is mandatory"
        if str(val).strip() not in VALID_COMPANY_CODES:
            return f"COMPANYCODE: '{str(val).strip()}' is invalid — must be one of 1001 / 1006 / 1009"
        return ""

    def get_rules(self) -> dict:
        return {
            "PLANT":          self.validate_plant,
            "NAME":           self.validate_name,
            "ADDRESS":        self.validate_address,
            "TCPL_PLANTTYPE": self.validate_tcpl_planttype,
            "COMPANYCODE":    self.validate_companycode,
        }


# ══════════════════════════════════════════════
#  Validator
# ══════════════════════════════════════════════
class SiteTableValidator:

    def __init__(self, site_path: str, part_path: str, valid_plants: list):
        self.site_path    = site_path
        self.part_path    = part_path
        self.valid_plants = valid_plants
        self.df           = pd.DataFrame()
        self.part_plants  = set()
        self.error_map    = {}
        self.reason_map   = {}

    def load(self):
        self.df = pd.read_excel(self.site_path, dtype=str)
        self.df.columns = [c.strip().upper() for c in self.df.columns]

        part_df = pd.read_excel(self.part_path, dtype=str)
        part_df.columns = [c.strip().upper() for c in part_df.columns]

        if "PLANT" not in part_df.columns:
            raise ValueError("PLANT column not found in Part table.")

        self.part_plants = set(part_df["PLANT"].dropna().str.strip().tolist())
        print(f"    Part table plants loaded  : {len(self.part_plants)} unique values")

    def validate(self):
        engine = SiteRuleEngine(self.valid_plants, self.part_plants)
        rules  = engine.get_rules()

        for idx, row in self.df.iterrows():
            failed_cols    = []
            col_reason_map = {}

            for col, rule_fn in rules.items():
                if col not in self.df.columns:
                    continue
                try:
                    reason = rule_fn(row)
                except Exception:
                    reason = f"{col}: Unexpected validation error"

                if reason:
                    failed_cols.append(col)
                    col_reason_map[col] = reason

            if failed_cols:
                self.error_map[idx]  = failed_cols
                self.reason_map[idx] = col_reason_map

    def get_error_series(self) -> pd.Series:
        result = {}
        for idx, col_reason in self.reason_map.items():
            result[idx] = " | ".join(col_reason.values())
        return pd.Series(result, dtype=str)

    def get_field_error_series(self, field_name: str) -> pd.Series:
        result = {}
        for idx, col_reason in self.reason_map.items():
            if field_name in col_reason:
                result[idx] = col_reason[field_name]
        return pd.Series(result, dtype=str)

    def get_errors_by_field(self) -> dict:
        field_errors: dict = {}
        for row_idx, bad_cols in self.error_map.items():
            for col in bad_cols:
                field_errors.setdefault(col, []).append(row_idx)
        return field_errors

    def get_plant_error_subcounts(self) -> dict:
        """
        Returns a dict with three keys:
          'blank'        – PLANT field was empty
          'not_in_pl'    – Plant absent from the Consolidated PL list
          'no_part_site' – Plant in PL list but no Part-Site combination
        """
        counts = {"blank": 0, "not_in_pl": 0, "no_part_site": 0}
        for idx, col_reason in self.reason_map.items():
            reason = col_reason.get("PLANT", "")
            if not reason:
                continue
            if "blank" in reason.lower():
                counts["blank"] += 1
            elif "consolidated pl list" in reason.lower():
                counts["not_in_pl"] += 1
            elif "part-site combination" in reason.lower():
                counts["no_part_site"] += 1
        return counts


# ══════════════════════════════════════════════
#  Report Writer
# ══════════════════════════════════════════════
class SiteReportWriter:

    # ── Sheet names ──────────────────────────
    # SHEET_ALL   = "Full Data"        # ← commented out: Full Data sheet disabled
    SHEET_SUMMARY = "Summary"
    SHEET_RULES   = "Rules"

    RULES_CONTENT = {
        "PLANT": [
            "Must not be blank.",
            "Must be present in the Consolidated PL list (hardcoded in the script).",
            "Must have an active Part-Site combination in the Part master table (PLANT column).",
        ],
        "NAME":           ["Must not be blank."],
        "ADDRESS":        ["Must not be blank."],
        "TCPL_PLANTTYPE": ["Must not be blank."],
        "COMPANYCODE": [
            "Must not be blank.",
            "Value must be one of: 1001 / 1006 / 1009.",
        ],
    }

    def __init__(self, validator: SiteTableValidator, output_path: str):
        self.validator   = validator
        self.output_path = output_path

    # ── helpers ──────────────────────────────
    def _write_header(self, ws, columns):
        for c_idx, col_name in enumerate(columns, start=1):
            cell           = ws.cell(row=1, column=c_idx, value=col_name)
            cell.fill      = HDR_FILL
            cell.font      = HDR_FONT
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border    = THIN_BORDER

    def _write_rows(self, ws, df: pd.DataFrame):
        for r_idx, (_, row) in enumerate(df.iterrows(), start=2):
            for c_idx, value in enumerate(row, start=1):
                cell      = ws.cell(row=r_idx, column=c_idx, value=value)
                cell.font = BODY_FONT

    def _highlight_full_data(self, ws, df: pd.DataFrame, error_map: dict, col_index: dict):
        error_row_set = set(error_map.keys())
        for df_idx in range(len(df)):
            row_fill = ROW_FILL if df_idx in error_row_set else WHITE_FILL
            for c in range(1, len(df.columns) + 1):
                ws.cell(row=df_idx + 2, column=c).fill = row_fill

        for df_idx, bad_cols in error_map.items():
            excel_row = df_idx + 2
            for col_name in bad_cols:
                if col_name in col_index:
                    cell      = ws.cell(row=excel_row, column=col_index[col_name])
                    cell.fill = RED_FILL
                    cell.font = ERR_FONT

    def _set_widths(self, ws):
        for col in ws.columns:
            max_len = max((len(str(c.value)) if c.value else 0) for c in col)
            ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 4, 60)

    def _style_summary_row(self, ws, row_num: int, num_cols: int = 7,
                           bold: bool = False, fill: PatternFill = None):
        for c in range(1, num_cols + 1):
            cell           = ws.cell(row=row_num, column=c)
            cell.font      = Font(name="Arial", bold=bold, size=10)
            cell.border    = THIN_BORDER
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if fill:
                cell.fill = fill

    # ══════════════════════════════════════════
    #  Summary sheet
    # ══════════════════════════════════════════
    def _write_summary_sheet(self, wb, error_map: dict, total_rows: int):
        ws = wb.create_sheet(self.SHEET_SUMMARY)

        # ── Title ──
        ws.merge_cells("A1:G1")
        title_cell           = ws.cell(row=1, column=1, value="Site Validation Summary")
        title_cell.font      = Font(name="Arial", bold=True, size=14)
        title_cell.fill      = TITLE_FILL
        title_cell.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[1].height = 24

        # ── Column headers ──
        headers = ["#", "Field Name", "Error Count", "Record Count", "% Health", "% of Error", "Reason / Sub-Category"]
        for c_idx, h in enumerate(headers, start=1):
            cell           = ws.cell(row=2, column=c_idx, value=h)
            cell.fill      = TITLE_FILL
            cell.font      = Font(name="Arial", bold=True)
            cell.border    = THIN_BORDER
            cell.alignment = Alignment(horizontal="center", vertical="center")

        # ── Build per-field error counts (keyed by field name) ──
        col_error_counts: dict = {}
        for bad_cols in error_map.values():
            for col in bad_cols:
                col_error_counts[col] = col_error_counts.get(col, 0) + 1

        # ── PLANT sub-buckets ──
        plant_subcounts = self.validator.get_plant_error_subcounts()

        # ── Iterate fields in canonical FIELD_ORDER (same as Rules sheet) ──
        row_num   = 3
        field_num = 1

        for col_name in FIELD_ORDER:
            count      = col_error_counts.get(col_name, 0)   # 0 if no errors for this field
            has_errors = count > 0

            pct_error  = round((count / total_rows) * 100, 2) if total_rows else 0
            pct_health = round(100 - pct_error, 2)

            # Choose fill: green tint for zero-error rows, white otherwise
            row_fill = NO_ERR_FILL if not has_errors else None

            ws.cell(row=row_num, column=1, value=field_num)
            ws.cell(row=row_num, column=2, value=col_name)
            ws.cell(row=row_num, column=3, value=count)
            ws.cell(row=row_num, column=4, value=total_rows)
            ws.cell(row=row_num, column=5, value=f"{pct_health}%")
            ws.cell(row=row_num, column=6, value=f"{pct_error}%")
            # Reason column: blank when no errors, blank for PLANT parent row (sub-rows carry reasons)
            ws.cell(row=row_num, column=7, value="" if (not has_errors or col_name == "PLANT") else "")

            self._style_summary_row(ws, row_num, fill=row_fill)
            row_num += 1

            # ── PLANT sub-rows (only when PLANT has errors) ──
            if col_name == "PLANT" and has_errors:
                sub_definitions = [
                    (
                        "  ↳ Not in Consolidated PL List",
                        plant_subcounts["blank"] + plant_subcounts["not_in_pl"],
                        "Blank plant code or plant absent from the Consolidated PL list",
                    ),
                    (
                        "  ↳ No Part-Site Combination",
                        plant_subcounts["no_part_site"],
                        "Plant is in PL list but has no active Part-Site record in Part master",
                    ),
                ]

                for sub_label, sub_count, sub_reason in sub_definitions:
                    sub_pct_err    = round((sub_count / total_rows) * 100, 2) if total_rows else 0
                    sub_pct_health = round(100 - sub_pct_err, 2)

                    ws.cell(row=row_num, column=1, value="")
                    ws.cell(row=row_num, column=2, value=sub_label)
                    ws.cell(row=row_num, column=3, value=sub_count)
                    ws.cell(row=row_num, column=4, value=total_rows)
                    ws.cell(row=row_num, column=5, value=f"{sub_pct_health}%")
                    ws.cell(row=row_num, column=6, value=f"{sub_pct_err}%")
                    ws.cell(row=row_num, column=7, value=sub_reason)

                    self._style_summary_row(ws, row_num, fill=PLANT_SUB_FILL)

                    ws.cell(row=row_num, column=2).alignment = Alignment(
                        horizontal="left", vertical="center", indent=1
                    )
                    ws.cell(row=row_num, column=7).alignment = Alignment(
                        horizontal="left", vertical="center", wrap_text=True
                    )
                    ws.cell(row=row_num, column=2).font = Font(name="Arial", size=10, italic=True)
                    ws.cell(row=row_num, column=7).font = Font(name="Arial", size=10, italic=True)

                    row_num += 1

            field_num += 1

        # ── TOTAL row (only counts fields that actually had errors) ──
        total_errors       = sum(col_error_counts.values())
        fields_with_errors = len(col_error_counts)
        total_record_count = total_rows * len(FIELD_ORDER)   # denominator = all fields × all rows
        total_pct_error    = round((total_errors / total_record_count) * 100, 2) if total_record_count else 0
        total_pct_health   = round(100 - total_pct_error, 2)

        ws.cell(row=row_num, column=1, value="")
        ws.cell(row=row_num, column=2, value="TOTAL")
        ws.cell(row=row_num, column=3, value=total_errors)
        ws.cell(row=row_num, column=4, value=total_record_count)
        ws.cell(row=row_num, column=5, value=f"{total_pct_health}%")
        ws.cell(row=row_num, column=6, value=f"{total_pct_error}%")
        ws.cell(row=row_num, column=7, value="")

        for c in range(1, 8):
            ws.cell(row=row_num, column=c).font      = Font(name="Arial", bold=True)
            ws.cell(row=row_num, column=c).fill      = TOTAL_FILL
            ws.cell(row=row_num, column=c).border    = THIN_BORDER
            ws.cell(row=row_num, column=c).alignment = Alignment(horizontal="center")

        row_num += 2   # blank spacer

        # ── Quick-glance stats block ──
        records_with_errors = len(error_map)
        records_passing     = total_rows - records_with_errors

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
            value_cell.font      = Font(name="Arial", size=10)
            value_cell.border    = THIN_BORDER
            value_cell.alignment = Alignment(horizontal="center", vertical="center")

            row_num += 1

        # ── Column widths ──
        col_widths = [6, 30, 14, 16, 12, 12, 65]
        for c_idx, width in enumerate(col_widths, start=1):
            ws.column_dimensions[get_column_letter(c_idx)].width = width

    # ── Per-field error sheets ────────────────
    def _write_field_error_sheets(self, wb, df: pd.DataFrame):
        field_errors = self.validator.get_errors_by_field()

        # ── Iterate in FIELD_ORDER so tab order matches Rules sheet ──
        for field_name in FIELD_ORDER:
            if field_name not in field_errors:
                continue   # no errors for this field — skip sheet creation

            row_indices = field_errors[field_name]
            sheet_name  = field_name[:31].replace("/", "-").replace("\\", "-").replace("*", "")
            ws          = wb.create_sheet(sheet_name)

            subset    = df.loc[row_indices].copy()
            keep_here = [c for c in KEEP_COLS if c in subset.columns] + ["ERROR_COLUMNS"]
            subset    = subset[keep_here]

            field_err_series        = self.validator.get_field_error_series(field_name)
            subset["ERROR_COLUMNS"] = subset.index.map(
                lambda i: field_err_series.get(i, "")
            )

            self._write_header(ws, subset.columns)
            col_idx_map = {col: i for i, col in enumerate(subset.columns, start=1)}

            for excel_row, (orig_idx, row_data) in enumerate(subset.iterrows(), start=2):
                for c_idx, (col, value) in enumerate(zip(subset.columns, row_data), start=1):
                    cell           = ws.cell(row=excel_row, column=c_idx, value=value)
                    cell.font      = BODY_FONT
                    cell.alignment = Alignment(vertical="center")
                    cell.fill      = ROW_FILL

                if field_name in col_idx_map:
                    target_cell      = ws.cell(row=excel_row, column=col_idx_map[field_name])
                    target_cell.fill = RED_FILL
                    target_cell.font = ERR_FONT

            self._set_widths(ws)

            note_row = len(subset) + 3
            ws.cell(
                row=note_row, column=1,
                value=f"Total error rows for '{field_name}': {len(subset)}",
            ).font = Font(name="Arial", italic=True, size=9, bold=True)

    # ── Rules sheet ───────────────────────────
    def _write_rules_sheet(self, wb):
        ws = wb.create_sheet(self.SHEET_RULES)

        ws.merge_cells("A1:C1")
        title_cell           = ws.cell(row=1, column=1, value="Site Table – Validation Rules")
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

        for field in FIELD_ORDER:                          # ← use canonical order
            rules_list = self.RULES_CONTENT.get(field, [])
            num_rules  = len(rules_list)

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
                field_cell.alignment = Alignment(vertical="center")

                desc_cell           = ws.cell(row=current_row, column=3, value=rule_text)
                desc_cell.font      = BODY_FONT
                desc_cell.border    = THIN_BORDER
                desc_cell.alignment = Alignment(wrap_text=True, vertical="center")

                current_row += 1

            if num_rules > 1:
                s = current_row - num_rules
                e = current_row - 1
                ws.merge_cells(start_row=s, start_column=1, end_row=e, end_column=1)
                ws.merge_cells(start_row=s, start_column=2, end_row=e, end_column=2)

            rule_num += 1

        ws.column_dimensions["A"].width = 6
        ws.column_dimensions["B"].width = 22
        ws.column_dimensions["C"].width = 70

    # ── Main write ────────────────────────────
    def write(self):
        v  = self.validator
        df = v.df.copy()

        error_series        = v.get_error_series()
        df["ERROR_COLUMNS"] = df.index.map(
            lambda i: error_series.get(i, "") if i in error_series.index else ""
        )

        keep_cols = [c for c in KEEP_COLS if c in df.columns] + ["ERROR_COLUMNS"]
        df        = df[keep_cols]

        col_index = {col: i for i, col in enumerate(df.columns, start=1)}

        # ── Create workbook (first sheet is the default active one) ──
        wb = Workbook()

        # ── Sheet order: Summary → Rules → error sheets (in FIELD_ORDER) ──

        # Summary  (rename the auto-created default sheet)
        ws_summary       = wb.active
        ws_summary.title = self.SHEET_SUMMARY
        self._write_summary_sheet_into(ws_summary, v.error_map, total_rows=len(df))

        # Rules
        self._write_rules_sheet(wb)

        # Per-field error sheets in FIELD_ORDER
        self._write_field_error_sheets(wb, df)

        # ── Full Data sheet is DISABLED ──
        # ws_all       = wb.create_sheet(self.SHEET_ALL)
        # ws_all.title = self.SHEET_ALL
        # self._write_header(ws_all, df.columns)
        # self._write_rows(ws_all, df)
        # self._highlight_full_data(ws_all, df, v.error_map, col_index)
        # self._set_widths(ws_all)
        # ws_all.freeze_panes = "A2"

        wb.save(self.output_path)
        print(f"\n✅  Output saved  → {self.output_path}")
        print(f"   Total rows    : {len(df)}")
        print(f"   Error rows    : {len(v.error_map)}")
        print(f"   Field sheets  : {[f for f in FIELD_ORDER if f in v.get_errors_by_field()]}")

    # ── Refactored: write summary content into an existing sheet ─────────
    def _write_summary_sheet_into(self, ws, error_map: dict, total_rows: int):
        """Same logic as _write_summary_sheet but writes into a pre-created ws."""

        # ── Title ──
        ws.merge_cells("A1:G1")
        title_cell           = ws.cell(row=1, column=1, value="Site Validation Summary")
        title_cell.font      = Font(name="Arial", bold=True, size=14)
        title_cell.fill      = TITLE_FILL
        title_cell.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[1].height = 24

        # ── Column headers ──
        headers = ["#", "Field Name", "Error Count", "Record Count", "% Health", "% of Error", "Reason / Sub-Category"]
        for c_idx, h in enumerate(headers, start=1):
            cell           = ws.cell(row=2, column=c_idx, value=h)
            cell.fill      = TITLE_FILL
            cell.font      = Font(name="Arial", bold=True)
            cell.border    = THIN_BORDER
            cell.alignment = Alignment(horizontal="center", vertical="center")

        col_error_counts: dict = {}
        for bad_cols in error_map.values():
            for col in bad_cols:
                col_error_counts[col] = col_error_counts.get(col, 0) + 1

        plant_subcounts = self.validator.get_plant_error_subcounts()

        row_num   = 3
        field_num = 1

        for col_name in FIELD_ORDER:
            count      = col_error_counts.get(col_name, 0)
            has_errors = count > 0

            pct_error  = round((count / total_rows) * 100, 2) if total_rows else 0
            pct_health = round(100 - pct_error, 2)

            row_fill = NO_ERR_FILL if not has_errors else None

            ws.cell(row=row_num, column=1, value=field_num)
            ws.cell(row=row_num, column=2, value=col_name)
            ws.cell(row=row_num, column=3, value=count)
            ws.cell(row=row_num, column=4, value=total_rows)
            ws.cell(row=row_num, column=5, value=f"{pct_health}%")
            ws.cell(row=row_num, column=6, value=f"{pct_error}%")
            ws.cell(row=row_num, column=7, value="")
            self._style_summary_row(ws, row_num, fill=row_fill)
            row_num += 1

            if col_name == "PLANT" and has_errors:
                sub_definitions = [
                    (
                        "  ↳ Not in Consolidated PL List",
                        plant_subcounts["blank"] + plant_subcounts["not_in_pl"],
                        "Blank plant code or plant absent from the Consolidated PL list",
                    ),
                    (
                        "  ↳ No Part-Site Combination",
                        plant_subcounts["no_part_site"],
                        "Plant is in PL list but has no active Part-Site record in Part master",
                    ),
                ]
                for sub_label, sub_count, sub_reason in sub_definitions:
                    sub_pct_err    = round((sub_count / total_rows) * 100, 2) if total_rows else 0
                    sub_pct_health = round(100 - sub_pct_err, 2)

                    ws.cell(row=row_num, column=1, value="")
                    ws.cell(row=row_num, column=2, value=sub_label)
                    ws.cell(row=row_num, column=3, value=sub_count)
                    ws.cell(row=row_num, column=4, value=total_rows)
                    ws.cell(row=row_num, column=5, value=f"{sub_pct_health}%")
                    ws.cell(row=row_num, column=6, value=f"{sub_pct_err}%")
                    ws.cell(row=row_num, column=7, value=sub_reason)

                    self._style_summary_row(ws, row_num, fill=PLANT_SUB_FILL)

                    ws.cell(row=row_num, column=2).alignment = Alignment(
                        horizontal="left", vertical="center", indent=1
                    )
                    ws.cell(row=row_num, column=7).alignment = Alignment(
                        horizontal="left", vertical="center", wrap_text=True
                    )
                    ws.cell(row=row_num, column=2).font = Font(name="Arial", size=10, italic=True)
                    ws.cell(row=row_num, column=7).font = Font(name="Arial", size=10, italic=True)

                    row_num += 1

            field_num += 1

        # ── TOTAL row ──
        total_errors       = sum(col_error_counts.values())
        total_record_count = total_rows * len(FIELD_ORDER)
        total_pct_error    = round((total_errors / total_record_count) * 100, 2) if total_record_count else 0
        total_pct_health   = round(100 - total_pct_error, 2)

        ws.cell(row=row_num, column=1, value="")
        ws.cell(row=row_num, column=2, value="TOTAL")
        ws.cell(row=row_num, column=3, value=total_errors)
        ws.cell(row=row_num, column=4, value=total_record_count)
        ws.cell(row=row_num, column=5, value=f"{total_pct_health}%")
        ws.cell(row=row_num, column=6, value=f"{total_pct_error}%")
        ws.cell(row=row_num, column=7, value="")

        for c in range(1, 8):
            ws.cell(row=row_num, column=c).font      = Font(name="Arial", bold=True)
            ws.cell(row=row_num, column=c).fill      = TOTAL_FILL
            ws.cell(row=row_num, column=c).border    = THIN_BORDER
            ws.cell(row=row_num, column=c).alignment = Alignment(horizontal="center")

        row_num += 2

        # ── Quick-glance stats block ──
        records_with_errors = len(error_map)
        records_passing     = total_rows - records_with_errors

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
            value_cell.font      = Font(name="Arial", size=10)
            value_cell.border    = THIN_BORDER
            value_cell.alignment = Alignment(horizontal="center", vertical="center")

            row_num += 1

        col_widths = [6, 30, 14, 16, 12, 12, 65]
        for c_idx, width in enumerate(col_widths, start=1):
            ws.column_dimensions[get_column_letter(c_idx)].width = width


# ══════════════════════════════════════════════
#  Orchestrator
# ══════════════════════════════════════════════
class SiteTableProcessor:

    def __init__(self, site_path: str, part_path: str, output_path: str, valid_plants: list):
        self.validator = SiteTableValidator(site_path, part_path, valid_plants)
        self.writer    = SiteReportWriter(self.validator, output_path)

    def run(self):
        print("📂  Loading files …")
        self.validator.load()
        print(f"    Site columns detected : {list(self.validator.df.columns)}")
        print("🔍  Validating rules …")
        self.validator.validate()
        print("📝  Writing report …")
        self.writer.write()


# ══════════════════════════════════════════════
#  Entry Point
# ══════════════════════════════════════════════
if __name__ == "__main__":
    processor = SiteTableProcessor(
        site_path    = SITE_INPUT_FILE,
        part_path    = PART_INPUT_FILE,
        output_path  = OUTPUT_FILE,
        valid_plants = VALID_PLANTS,
    )
    processor.run()
