# import pandas as pd
# from openpyxl import Workbook
# from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
# from openpyxl.utils import get_column_letter

# # ─────────────────────────────────────────────
# #  FILE PATHS
# # ─────────────────────────────────────────────
# INPUT_FILE  = r"C:\Users\M sD\Downloads\Data_rulesets_check\Excel_Files\Part_site_FG 10.04.2026.csv"
# OUTPUT_FILE = r"C:\Users\M sD\Downloads\Data_rulesets_check\Output_Files\Validated_Part.xlsx"

# # ─────────────────────────────────────────────
# #  CONSOLIDATED PL LIST
# # ─────────────────────────────────────────────
# VALID_PLANTS = {
#    "1127", "1100", "1105", "1146", "1156", "1107", "1157", "1158", "1166", "1180",
#     "1184", "1186", "1197", "1203", "1204", "1211", "1213", "1214", "1218", "1223",
#     "1110", "1225", "1226", "1229", "1233", "1234", "1240", "1113", "1248", "1253",
#     "1257", "1258", "1265", "1114", "1145", "1275", "1279", "1416", "1421", "1423",
#     "1425", "1426", "1428", "1429", "1430", "1432", "1433", "1436", "1437", "1438",
#     "1439", "1440", "1442", "1445", "1445", "1449", "1451", "1452", "1455", "1463",
#     "1471", "1473", "1475", "1476", "1477", "1478", "1480", "1481", "1483", "1484",
#     "1485", "1487", "1488", "1233", "1491", "1495", "1500", "1501", "1505", "1233",
#     "1506", "1507", "1421", "1508", "1509", "1521", "1525", "1563", "1578", "1650",
#     "1651", "1652", "1654", "1656", "1657", "1658", "1659", "1661", "1579", "1627",
#     "1589", "1623", "1509", "1646", "1647", "1640", "1112", "5011637", "5123296",
#     "5123742", "4007430", "1623", "1722", "1724", "1725", "1726", "1731", "1732",
#     "1733", "1734", "1738", "1739", "1740", "1742", "1754", "1757", "1758", "1771",
#     "1774", "1780", "1784", "1785", "1788", "1511", "1512", "1448", "4011702",
#     "5013796", "5018849", "5011407", "5015073", "5011308", "5123531", "1642", "1638",
#     "1104", "1104A", "1643", "1642", "1106", "1647", "1643", "1109", "1645",
#     "1645", "1111", "1648", "1648", "1649", "1649", "1653", "1653", "1801", "1802",
#     "2082", "2091", "2088", "2089", "2083", "2084", "2085", "2090", "2081",
#     "2086", "2087", "1554", "1520", "1472", "1571", "1298", "1295", "1196",
#     "1292", "1176", "1441", "1522", "1494", "1489", "1518", "1249", "1296",
#     "1137", "1208", "1155", "1569", "1235", "1281", "1503", "1482", "1135",
#     "1205", "1241", "1499", "1462", "1555", "1559", "1575", "2092", "1558",
#     "1601", "1125", "1256", "1568"
# }

# # ─────────────────────────────────────────────
# #  STYLING CONSTANTS
# # ─────────────────────────────────────────────
# RED_FILL   = PatternFill("start_color", "FF0000", "FF0000", fill_type="solid")
# ROW_FILL   = PatternFill("start_color", "FFF2CC", "FFF2CC", fill_type="solid")
# HDR_FILL   = PatternFill("start_color", "D9E1F2", "D9E1F2", fill_type="solid")
# RULE_FILL  = PatternFill("start_color", "E2EFDA", "E2EFDA", fill_type="solid")
# TITLE_FILL = PatternFill("start_color", "BDD7EE", "BDD7EE", fill_type="solid")
# TOTAL_FILL = PatternFill("start_color", "F2F2F2", "F2F2F2", fill_type="solid")
# STATS_FILL = PatternFill("start_color", "EDEDED", "EDEDED", fill_type="solid")

# HDR_FONT    = Font(bold=True, name="Arial")
# BODY_FONT   = Font(name="Arial", size=10)
# ERR_FONT    = Font(name="Arial", size=10, bold=True, color="FFFFFF")

# THIN_BORDER = Border(
#     left=Side(style="thin"), right=Side(style="thin"),
#     top=Side(style="thin"),  bottom=Side(style="thin"),
# )


# # ══════════════════════════════════════════════
# #  Rule Engine
# # ══════════════════════════════════════════════
# class RuleEngine:

#     def __init__(self, valid_plants: list):
#         self.valid_plants = set(valid_plants)
        
#     @staticmethod
#     def _is_blank(value) -> bool:
#         if pd.isna(value):
#             return True
#         return str(value).strip() == ""

#     # REASON constants extracted for tuple outputs
#     M_NUMBER_REASON = "MATERIALNUMBER: Must be 14xxxxxxxxxxxxxx (FERT) or 15xxxxxxxxxxxxxx (HAWA) — invalid range or blank"
#     PLANT_REASON    = "PLANT: Value is not in the Consolidated PL list"
#     PRODDESC_REASON = "PRODUCTDESCRIPTION: Field is blank — description is mandatory"
#     PRODTYPE_REASON = "PRODUCTTYPE: Must be FERT or HAWA — invalid or blank value found"
#     PRODHIER_REASON = "PRODUCTHIERARCHY: Field is blank — hierarchy is mandatory"
#     MRPTYPE_REASON  = "MRPTYPE: Field is blank"
#     PROC_REASON     = "PROCUREMENTTYPE: Field is blank"
#     ABC_REASON      = "ABCINDICATOR: Field is blank — defaulting to 100% error"
#     IBP_REASON      = "IBPSTATUS: Must be 'IBP' or blank — unexpected value found"
#     XPLANT_REASON   = "XPLANTMATSTATUS: Must be '2' or blank — unexpected value found"

#     def validate_material_number(self, row) -> tuple[bool, str]:
#         raw = row.get("MATERIALNUMBER")
#         if self._is_blank(raw):
#             return False, self.M_NUMBER_REASON
#         valstr = str(raw).strip()
#         if not valstr.isdigit():
#             return False, self.M_NUMBER_REASON
        
#         val = int(valstr)
#         mat_type = str(row.get("PRODUCTTYPE", "")).strip().upper()

#         if mat_type == "FERT":
#             if 14000000000000 <= val <= 14999999999999:
#                 return True, ""
#         elif mat_type == "HAWA":
#             if 15000000000000 <= val <= 15999999999999:
#                 return True, ""
#         return False, self.M_NUMBER_REASON

#     def validate_plant(self, row) -> tuple[bool, str]:
#         val = str(row.get("PLANT", "")).strip().upper()
#         if self._is_blank(val) or val not in self.valid_plants:
#             return False, self.PLANT_REASON
#         return True, ""

#     def validate_product_description(self, row) -> tuple[bool, str]:
#         if self._is_blank(row.get("PRODUCTDESCRIPTION")):
#             return False, self.PRODDESC_REASON
#         return True, ""

#     def validate_product_type(self, row) -> tuple[bool, str]:
#         val = str(row.get("PRODUCTTYPE", "")).strip().upper()
#         if val in {"FERT", "HAWA"}:
#             return True, ""
#         return False, self.PRODTYPE_REASON

#     def validate_product_hierarchy(self, row) -> tuple[bool, str]:
#         if self._is_blank(row.get("PRODUCTHIERARCHY")):
#             return False, self.PRODHIER_REASON
#         return True, ""

#     def validate_mrp_type(self, row) -> tuple[bool, str]:
#         val = str(row.get("MRPTYPE", "")).strip().upper()
#         if val in {"ND", "PD"}:
#             return True, ""
#         return False, self.MRPTYPE_REASON

#     def validate_procurement_type(self, row) -> tuple[bool, str]:
#         if self._is_blank(row.get("PROCUREMENTTYPE")):
#             return False, self.PROC_REASON
#         return True, ""

#     def validate_abc_indicator(self, row) -> tuple[bool, str]:
#         # Force 100% error as requested
#         return False, self.ABC_REASON

#     def validate_ibp_status(self, row) -> tuple[bool, str]:
#         raw = row.get("IBPSTATUS")
#         if self._is_blank(raw):
#             return True, ""
#         if str(raw).strip().upper() == "IBP":
#             return True, ""
#         return False, self.IBP_REASON

#     def validate_xplant_mat_status(self, row) -> tuple[bool, str]:
#         raw = row.get("XPLANTMATSTATUS")
#         if self._is_blank(raw):
#             return True, ""
#         if str(raw).strip() in {"2", "02"}:
#             return True, ""
#         return False, self.XPLANT_REASON

#     def get_rules(self) -> dict:
#         return {
#             "MATERIALNUMBER":     self.validate_material_number,
#             "PLANT":              self.validate_plant,
#             "PRODUCTDESCRIPTION": self.validate_product_description,
#             "PRODUCTTYPE":        self.validate_product_type,
#             "PRODUCTHIERARCHY":   self.validate_product_hierarchy,
#             "MRPTYPE":            self.validate_mrp_type,
#             "PROCUREMENTTYPE":    self.validate_procurement_type,
#             "ABCINDICATOR":       self.validate_abc_indicator,
#             "IBPSTATUS":          self.validate_ibp_status,
#             "XPLANTMATSTATUS":    self.validate_xplant_mat_status,
#         }


# # ══════════════════════════════════════════════
# #  Validator
# # ══════════════════════════════════════════════
# class PartTableValidator:

#     def __init__(self, filepath: str, valid_plants: list):
#         self.filepath     = filepath
#         self.valid_plants = valid_plants
#         self.df           = pd.DataFrame()
#         self.error_map    = {}

#     def load(self):
#         if self.filepath.lower().endswith('.csv'):
#             self.df = pd.read_csv(self.filepath, dtype=str)
#         else:
#             self.df = pd.read_excel(self.filepath, dtype=str)
            
#         self.df.columns = [c.strip().upper() for c in self.df.columns]

#     def validate(self):
#         engine = RuleEngine(self.valid_plants)
#         rules  = engine.get_rules()

#         for idx, row in self.df.iterrows():
#             errors = {}
#             for col, rule_fn in rules.items():
#                 if col not in self.df.columns and col not in ("ABCINDICATOR"):
#                     # Abc indicator might not be in columns physically but we force fail it anyway 
#                     if col != "ABCINDICATOR": 
#                         continue
                
#                 try:
#                     passed, reason = rule_fn(row)
#                 except Exception as e:
#                     passed, reason = False, f"Exception: {str(e)}"

#                 if not passed:
#                     errors[col] = reason

#             if errors:
#                 self.error_map[idx] = errors

#     def get_error_series(self, field_name: str) -> pd.Series:
#         result = {}
#         for idx, errdict in self.error_map.items():
#             if field_name in errdict:
#                 result[idx] = errdict[field_name]
#         return pd.Series(result, dtype=str)


# # ══════════════════════════════════════════════
# #  Report Writer
# # ══════════════════════════════════════════════
# class ReportWriter:

#     SHEET_SUMMARY = "Summary"
#     SHEET_RULES   = "Rules"

#     RULES_CONTENT = {
#         "MATERIALNUMBER": [
#             "Must not be blank.",
#             "For FERT type: Material number must be in range 14000000000000 – 14999999999999.",
#             "For HAWA type: Material number must be in range 15000000000000 – 15999999999999.",
#         ],
#         "PLANT": [
#             "Must not be blank.",
#             "Must be present in the Consolidated PL list (hardcoded in script).",
#         ],
#         "PRODUCTDESCRIPTION": ["Must not be blank."],
#         "PRODUCTTYPE": [
#             "Must not be blank.",
#             "Value must be either FERT or HAWA.",
#         ],
#         "PRODUCTHIERARCHY": ["Must not be blank."],
#         "MRPTYPE": [
#             "Must not be blank.",
#             "Value must be either ND or PD.",
#         ],
#         "PROCUREMENTTYPE": ["Must not be blank."],
#         "ABCINDICATOR": [
#             "Must not be blank.",
#             "NOTE: Column is fully blank in source data – defaulting to 100% error.",
#         ],
#         "IBPSTATUS": [
#             "Allowed values: IBP or blank.",
#             "Any other value is treated as an error.",
#         ],
#         "XPLANTMATSTATUS": [
#             "Allowed values: 2 or blank.",
#             "Any other value is treated as an error.",
#         ],
#     }

#     def __init__(self, validator: PartTableValidator, output_path: str):
#         self.validator   = validator
#         self.output_path = output_path

#     # ── Helpers ──────────────────────────────
#     def _write_header(self, ws, columns):
#         for c_idx, col_name in enumerate(columns, start=1):
#             cell           = ws.cell(row=1, column=c_idx, value=col_name)
#             cell.fill      = HDR_FILL
#             cell.font      = HDR_FONT
#             cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
#             cell.border    = THIN_BORDER
#         ws.row_dimensions[1].height = 30

#     def _auto_width(self, ws, min_w=10, max_w=60):
#         for col in ws.columns:
#             length = max((len(str(c.value)) if c.value else 0) for c in col)
#             ws.column_dimensions[get_column_letter(col[0].column)].width = min(max(length + 3, min_w), max_w)

#     # ── Summary sheet ────────────────────────
#     def _write_summary_sheet(self, wb, error_map: dict, total_rows: int):
#         ws = wb.create_sheet(self.SHEET_SUMMARY)

#         # Build list of all fields
#         ruleset_field_order = [
#             "MATERIALNUMBER", "PLANT", "PRODUCTDESCRIPTION", "PRODUCTTYPE",
#             "PRODUCTHIERARCHY", "MRPTYPE", "PROCUREMENTTYPE", "ABCINDICATOR", 
#             "IBPSTATUS", "XPLANTMATSTATUS"
#         ]

#         col_error_counts = {col: 0 for col in ruleset_field_order}
#         for bad_cols in error_map.values():
#             for col in bad_cols.keys():
#                 if col in col_error_counts:
#                     col_error_counts[col] += 1
#                 else:
#                     col_error_counts[col] = 1

#         REASON_MAP = {
#             "MATERIALNUMBER": "MATERIALNUMBER: Must be 14xxxxxxxxxxxxxx (FERT) or 15xxxxxxxxxxxxxx (HAWA) — invalid range or blank",
#             "PLANT": "PLANT: Value is not in the Consolidated PL list",
#             "PRODUCTDESCRIPTION": "PRODUCTDESCRIPTION: Field is blank",
#             "PRODUCTTYPE": "PRODUCTTYPE: Must be FERT or HAWA — invalid or blank value found",
#             "PRODUCTHIERARCHY": "PRODUCTHIERARCHY: Field is blank",
#             "MRPTYPE": "MRPTYPE: Field is blank",
#             "PROCUREMENTTYPE": "PROCUREMENTTYPE: Field is blank",
#             "ABCINDICATOR": "ABCINDICATOR: Field is blank",
#             "IBPSTATUS": "IBPSTATUS: Must be 'IBP' or blank — unexpected value found",
#             "XPLANTMATSTATUS": "XPLANTMATSTATUS: Must be '2' or blank — unexpected value found",
#         }

#         # Title
#         ws.merge_cells("A1:G1")
#         title_cell           = ws.cell(row=1, column=1, value="Part Master FG Validation Summary")
#         title_cell.font      = Font(name="Arial", bold=True, size=14)
#         title_cell.fill      = TITLE_FILL
#         title_cell.alignment = Alignment(horizontal="left", vertical="center")
#         ws.row_dimensions[1].height = 24

#         # Headers
#         headers = ["#", "Field Name", "Error Count", "Record Count", "% Health", "% of Error", "Reason"]
#         for c_idx, h in enumerate(headers, start=1):
#             cell           = ws.cell(row=2, column=c_idx, value=h)
#             cell.fill      = TITLE_FILL
#             cell.font      = Font(name="Arial", bold=True)
#             cell.border    = THIN_BORDER
#             cell.alignment = Alignment(horizontal="center", vertical="center")

#         # Per-field rows
#         row_num = 3
#         for field_num, col_name in enumerate(ruleset_field_order, start=1):
#             count = col_error_counts.get(col_name, 0)
#             pct_error  = round((count / total_rows) * 100, 2) if total_rows else 0
#             pct_health = round(100 - pct_error, 2)

#             reason_text = REASON_MAP.get(col_name, "Validation rule failed") if count > 0 else ""

#             ws.cell(row=row_num, column=1, value=field_num)
#             ws.cell(row=row_num, column=2, value=col_name)
#             ws.cell(row=row_num, column=3, value=count)
#             ws.cell(row=row_num, column=4, value=total_rows)
#             ws.cell(row=row_num, column=5, value=f"{pct_health}%")
#             ws.cell(row=row_num, column=6, value=f"{pct_error}%")
#             ws.cell(row=row_num, column=7, value=reason_text)

#             for c in range(1, 8):
#                 ws.cell(row=row_num, column=c).font      = BODY_FONT
#                 ws.cell(row=row_num, column=c).border    = THIN_BORDER
#                 align = Alignment(horizontal="center", vertical="center")
#                 if c == 7:
#                     align = Alignment(horizontal="left", vertical="center")
#                 ws.cell(row=row_num, column=c).alignment = align

#             row_num += 1

#         # TOTAL row
#         total_errors       = sum(col_error_counts.values())
#         total_record_count = total_rows * len(ruleset_field_order)
#         total_pct_error    = round((total_errors / total_record_count) * 100, 2) if total_record_count else 0
#         total_pct_health   = round(100 - total_pct_error, 2)

#         ws.cell(row=row_num, column=1, value="")
#         ws.cell(row=row_num, column=2, value="TOTAL")
#         ws.cell(row=row_num, column=3, value=total_errors)
#         ws.cell(row=row_num, column=4, value=total_record_count)
#         ws.cell(row=row_num, column=5, value=f"{total_pct_health}%")
#         ws.cell(row=row_num, column=6, value=f"{total_pct_error}%")
#         ws.cell(row=row_num, column=7, value="")

#         for c in range(1, 8):
#             ws.cell(row=row_num, column=c).font      = Font(name="Arial", bold=True)
#             ws.cell(row=row_num, column=c).fill      = TOTAL_FILL
#             ws.cell(row=row_num, column=c).border    = THIN_BORDER
#             ws.cell(row=row_num, column=c).alignment = Alignment(horizontal="center", vertical="center")

#         row_num += 2

#         # Stats Block
#         records_with_errors = len(error_map)
#         records_passing     = total_rows - records_with_errors

#         for label, value in [
#             ("Total Records:",       total_rows),
#             ("Records with Errors:", records_with_errors),
#             ("Records Passing:",     records_passing),
#         ]:
#             ws.merge_cells(start_row=row_num, start_column=1, end_row=row_num, end_column=2)
#             label_cell           = ws.cell(row=row_num, column=1, value=label)
#             label_cell.font      = Font(name="Arial", bold=True, size=10)
#             label_cell.fill      = STATS_FILL
#             label_cell.border    = THIN_BORDER
#             label_cell.alignment = Alignment(horizontal="left", vertical="center")

#             value_cell           = ws.cell(row=row_num, column=3, value=value)
#             value_cell.font      = BODY_FONT
#             value_cell.border    = THIN_BORDER
#             value_cell.alignment = Alignment(horizontal="center", vertical="center")

#             row_num += 1

#         self._auto_width(ws, min_w=8, max_w=70)

#     # ── Rule_Set sheet ────────────────────────
#     def _write_ruleset_sheet(self, wb):
#         ws = wb.create_sheet(self.SHEET_RULES)

#         ws.merge_cells("A1:C1")
#         title_cell           = ws.cell(row=1, column=1, value="Part Master FG – Validation Rules")
#         title_cell.font      = Font(name="Arial", bold=True, size=13)
#         title_cell.fill      = TITLE_FILL
#         title_cell.alignment = Alignment(horizontal="center")
#         ws.row_dimensions[1].height = 22

#         for c_idx, h in enumerate(["#", "Field", "Rule Description"], start=1):
#             cell           = ws.cell(row=3, column=c_idx, value=h)
#             cell.fill      = HDR_FILL
#             cell.font      = HDR_FONT
#             cell.border    = THIN_BORDER
#             cell.alignment = Alignment(horizontal="center")

#         current_row = 4
#         rule_num    = 1

#         for field, rules_list in self.RULES_CONTENT.items():
#             num_rules = len(rules_list)

#             for r_idx, rule_text in enumerate(rules_list):
#                 num_cell           = ws.cell(row=current_row, column=1, value=rule_num if r_idx == 0 else "")
#                 num_cell.font      = Font(name="Arial", size=10, bold=(r_idx == 0))
#                 num_cell.fill      = RULE_FILL
#                 num_cell.border    = THIN_BORDER
#                 num_cell.alignment = Alignment(horizontal="center", vertical="center")

#                 field_cell           = ws.cell(row=current_row, column=2, value=field if r_idx == 0 else "")
#                 field_cell.font      = Font(name="Arial", size=10, bold=(r_idx == 0))
#                 field_cell.fill      = RULE_FILL
#                 field_cell.border    = THIN_BORDER
#                 field_cell.alignment = Alignment(horizontal="center", vertical="center")

#                 desc_cell           = ws.cell(row=current_row, column=3, value=rule_text)
#                 desc_cell.font      = BODY_FONT
#                 desc_cell.border    = THIN_BORDER
#                 desc_cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")

#                 current_row += 1

#             if num_rules > 1:
#                 s = current_row - num_rules
#                 e = current_row - 1
#                 ws.merge_cells(start_row=s, start_column=1, end_row=e, end_column=1)
#                 ws.merge_cells(start_row=s, start_column=2, end_row=e, end_column=2)

#             rule_num += 1

#         ws.column_dimensions["A"].width = 6
#         ws.column_dimensions["B"].width = 30
#         ws.column_dimensions["C"].width = 65

#     # ── Field Error Sheets ───────────────────
#     def _write_field_error_sheets(self, wb, df: pd.DataFrame):
#         v = self.validator
        
#         all_error_fields = set()
#         for bad_cols in v.error_map.values():
#             all_error_fields.update(bad_cols.keys())
            
#         for datalake_field_name in sorted(all_error_fields):
#             # Skip creating specific error sheet for ABCINDICATOR as requested
#             if datalake_field_name == "ABCINDICATOR":
#                 continue

#             row_indices = [idx for idx, errdict in v.error_map.items() if datalake_field_name in errdict]
#             if not row_indices:
#                 continue
                
#             sheet_name = datalake_field_name[:31]
#             existing   = [s.title for s in wb.worksheets]
#             counter    = 1
#             base_name  = sheet_name
#             while sheet_name in existing:
#                 sheet_name = f"{base_name[:28]}_{counter}"
#                 counter   += 1

#             ws = wb.create_sheet(sheet_name)

#             subset = df.loc[row_indices].copy()

#             subset["ERROR_COLUMNS"] = subset.index.map(
#                 lambda i: v.error_map.get(i, {}).get(datalake_field_name, "")
#             )

#             self._write_header(ws, subset.columns)
#             col_idx_map = {col: i for i, col in enumerate(subset.columns, start=1)}
            
#             for r_idx, (orig_idx, row_data) in enumerate(subset.iterrows(), start=2):
#                 for c_idx, (col, value) in enumerate(zip(subset.columns, row_data), start=1):
#                     cell           = ws.cell(row=r_idx, column=c_idx, value=value)
#                     cell.font      = BODY_FONT
#                     cell.border    = THIN_BORDER
#                     cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
#                     cell.fill      = ROW_FILL
                    
#                 if datalake_field_name in col_idx_map:
#                     target_cell      = ws.cell(row=r_idx, column=col_idx_map[datalake_field_name])
#                     target_cell.fill = RED_FILL
#                     target_cell.font = ERR_FONT
                    
#             self._auto_width(ws, min_w=10, max_w=60)
#             ws.freeze_panes = "A2"
            
#             note_row = len(subset) + 3
#             ws.cell(
#                 row=note_row, column=1,
#                 value=f"Total error rows for '{datalake_field_name}': {len(subset)}",
#             ).font = Font(name="Arial", italic=True, size=9, bold=True)

#     # ── Orchestrate Writing ──────────────────
#     def write(self):
#         v  = self.validator
#         df = v.df.copy()

#         wb = Workbook()
#         if "Sheet" in wb.sheetnames:
#             del wb["Sheet"]

#         # Note: 'Full Data' is dropped in this structure per unified design.

#         # 1. Summary
#         self._write_summary_sheet(wb, v.error_map, total_rows=len(df))

#         # 2. Rules
#         self._write_ruleset_sheet(wb)

#         # 3. Field Errors
#         self._write_field_error_sheets(wb, df)

#         wb.save(self.output_path)
#         print(f"\n✅  Output saved  → {self.output_path}")
#         print(f"   Total rows    : {len(df)}")
#         print(f"   Error rows    : {len(v.error_map)}")


# # ══════════════════════════════════════════════
# #  Orchestrator
# # ══════════════════════════════════════════════
# class PartTableProcessor:

#     def __init__(self, input_path: str, output_path: str):
#         self.validator = PartTableValidator(input_path, list(VALID_PLANTS))
#         self.writer    = ReportWriter(self.validator, output_path)

#     def run(self):
#         print("📂  Loading file …")
#         self.validator.load()
#         print(f"    Columns detected : {list(self.validator.df.columns)}")
#         print("🔍  Validating rules …")
#         self.validator.validate()
#         print("📝  Writing report …")
#         self.writer.write()


# # ══════════════════════════════════════════════
# #  ENTRY POINT
# # ══════════════════════════════════════════════
# if __name__ == "__main__":
#     processor = PartTableProcessor(INPUT_FILE, OUTPUT_FILE)
#     processor.run()


import io
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────────
#  SHAREPOINT CREDENTIALS & CONFIG
# ─────────────────────────────────────────────
SHAREPOINT_URL      = "https://<your-tenant>.sharepoint.com"          # ← update
SHAREPOINT_SITE     = "/sites/<your-site>"                             # ← update
SHAREPOINT_FILE_PATH = "Shared Documents/<folder>/ConsolidatedPL.xlsx" # ← update relative path on SharePoint
SHAREPOINT_EMAIL    = "your.email@company.com"                         # ← update
SHAREPOINT_PASSWORD = "your_password"                                  # ← update
SHAREPOINT_SHEET    = "March 2006"                                     # subsheet name
SHAREPOINT_COL      = "Plant code"                                     # column name in that sheet

# ─────────────────────────────────────────────
#  FILE PATHS
# ─────────────────────────────────────────────
INPUT_FILE  = r"C:\Users\M sD\Downloads\Data_rulesets_check\Excel_Files\Part_site_FG 10.04.2026.csv"
OUTPUT_FILE = r"C:\Users\M sD\Downloads\Data_rulesets_check\Output_Files\Validated_Part.xlsx"

# ─────────────────────────────────────────────
#  LOAD VALID PLANTS FROM SHAREPOINT
# ─────────────────────────────────────────────
def load_valid_plants_from_sharepoint(
    sharepoint_url: str,
    sharepoint_site: str,
    file_path: str,
    email: str,
    password: str,
    sheet_name: str,
    column_name: str,
) -> set:
    """
    Connects to SharePoint with email/password credentials,
    downloads the Excel file, reads the given sheet, and returns
    a set of plant codes from the specified column.

    Requirements:
        pip install Office365-REST-Python-Client openpyxl
    """
    try:
        from office365.sharepoint.client_context import ClientContext
        from office365.runtime.auth.user_credential import UserCredential
    except ImportError:
        raise ImportError(
            "Office365-REST-Python-Client is not installed.\n"
            "Run: pip install Office365-REST-Python-Client"
        )

    print(f"🔐  Connecting to SharePoint as {email} …")
    ctx = ClientContext(sharepoint_url + sharepoint_site).with_credentials(
        UserCredential(email, password)
    )

    # Download file bytes
    file_url = f"{sharepoint_site}/{file_path}"
    response = ctx.web.get_file_by_server_relative_url(file_url).download().execute_query()

    print(f"📥  Downloading plant list from SharePoint …")
    # response.value holds the raw bytes of the file
    file_bytes = response.value

    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, dtype=str)
    df.columns = [c.strip() for c in df.columns]

    if column_name not in df.columns:
        raise ValueError(
            f"Column '{column_name}' not found in sheet '{sheet_name}'.\n"
            f"Available columns: {list(df.columns)}"
        )

    plants = set(
        str(v).strip().upper()
        for v in df[column_name].dropna()
        if str(v).strip() != ""
    )
    print(f"✅  Loaded {len(plants)} plant codes from SharePoint.")
    return plants


# ─────────────────────────────────────────────
#  STYLING CONSTANTS
# ─────────────────────────────────────────────
RED_FILL   = PatternFill("solid", fgColor="FF0000")
ROW_FILL   = PatternFill("solid", fgColor="FFF2CC")
HDR_FILL   = PatternFill("solid", fgColor="D9E1F2")
RULE_FILL  = PatternFill("solid", fgColor="E2EFDA")
TITLE_FILL = PatternFill("solid", fgColor="BDD7EE")
TOTAL_FILL = PatternFill("solid", fgColor="F2F2F2")
STATS_FILL = PatternFill("solid", fgColor="EDEDED")

HDR_FONT  = Font(bold=True, name="Arial")
BODY_FONT = Font(name="Arial", size=10)
ERR_FONT  = Font(name="Arial", size=10, bold=True, color="FFFFFF")

THIN_BORDER = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"),  bottom=Side(style="thin"),
)

# ─────────────────────────────────────────────
#  Columns shown in error sheets
#  (must match keys in RULES_CONTENT exactly)
# ─────────────────────────────────────────────
RULES_FIELDS_ORDERED = [
    "MATERIALNUMBER",
    "PLANT",
    "PRODUCTDESCRIPTION",
    "PRODUCTTYPE",
    "PRODUCTHIERARCHY",
    "MRPTYPE",
    "PROCUREMENTTYPE",
    "ABCINDICATOR",
    "IBPSTATUS",
    "XPLANTMATSTATUS",
]

# Priority order for error sub-sheets (change 2 → sheets for these 4 first)
ERROR_SHEET_PRIORITY = [
    "PLANT",
    "PROCUREMENTTYPE",
    "PRODUCTHIERARCHY",
    "MRPTYPE",
]


# ══════════════════════════════════════════════
#  Rule Engine
# ══════════════════════════════════════════════
class RuleEngine:

    def __init__(self, valid_plants: set):
        self.valid_plants = valid_plants

    @staticmethod
    def _is_blank(value) -> bool:
        if pd.isna(value):
            return True
        return str(value).strip() == ""

    M_NUMBER_REASON = "MATERIALNUMBER: Must be 14xxxxxxxxxxxxxx (FERT) or 15xxxxxxxxxxxxxx (HAWA) — invalid range or blank"
    PLANT_REASON    = "PLANT: Value is not in the Consolidated PL list"
    PRODDESC_REASON = "PRODUCTDESCRIPTION: Field is blank — description is mandatory"
    PRODTYPE_REASON = "PRODUCTTYPE: Must be FERT or HAWA — invalid or blank value found"
    PRODHIER_REASON = "PRODUCTHIERARCHY: Field is blank — hierarchy is mandatory"
    MRPTYPE_REASON  = "MRPTYPE: Field is blank"
    PROC_REASON     = "PROCUREMENTTYPE: Field is blank"
    ABC_REASON      = "ABCINDICATOR: Field is blank — defaulting to 100% error"
    IBP_REASON      = "IBPSTATUS: Must be 'IBP' or blank — unexpected value found"
    XPLANT_REASON   = "XPLANTMATSTATUS: Must be '2' or blank — unexpected value found"

    def validate_material_number(self, row) -> tuple[bool, str]:
        raw = row.get("MATERIALNUMBER")
        if self._is_blank(raw):
            return False, self.M_NUMBER_REASON
        valstr = str(raw).strip()
        if not valstr.isdigit():
            return False, self.M_NUMBER_REASON
        val = int(valstr)
        mat_type = str(row.get("PRODUCTTYPE", "")).strip().upper()
        if mat_type == "FERT" and 14000000000000 <= val <= 14999999999999:
            return True, ""
        if mat_type == "HAWA" and 15000000000000 <= val <= 15999999999999:
            return True, ""
        return False, self.M_NUMBER_REASON

    def validate_plant(self, row) -> tuple[bool, str]:
        val = str(row.get("PLANT", "")).strip().upper()
        if self._is_blank(val) or val not in self.valid_plants:
            return False, self.PLANT_REASON
        return True, ""

    def validate_product_description(self, row) -> tuple[bool, str]:
        if self._is_blank(row.get("PRODUCTDESCRIPTION")):
            return False, self.PRODDESC_REASON
        return True, ""

    def validate_product_type(self, row) -> tuple[bool, str]:
        val = str(row.get("PRODUCTTYPE", "")).strip().upper()
        if val in {"FERT", "HAWA"}:
            return True, ""
        return False, self.PRODTYPE_REASON

    def validate_product_hierarchy(self, row) -> tuple[bool, str]:
        if self._is_blank(row.get("PRODUCTHIERARCHY")):
            return False, self.PRODHIER_REASON
        return True, ""

    def validate_mrp_type(self, row) -> tuple[bool, str]:
        val = str(row.get("MRPTYPE", "")).strip().upper()
        if val in {"ND", "PD"}:
            return True, ""
        return False, self.MRPTYPE_REASON

    def validate_procurement_type(self, row) -> tuple[bool, str]:
        if self._is_blank(row.get("PROCUREMENTTYPE")):
            return False, self.PROC_REASON
        return True, ""

    def validate_abc_indicator(self, row) -> tuple[bool, str]:
        # Force 100% error as requested
        return False, self.ABC_REASON

    def validate_ibp_status(self, row) -> tuple[bool, str]:
        raw = row.get("IBPSTATUS")
        if self._is_blank(raw):
            return True, ""
        if str(raw).strip().upper() == "IBP":
            return True, ""
        return False, self.IBP_REASON

    def validate_xplant_mat_status(self, row) -> tuple[bool, str]:
        raw = row.get("XPLANTMATSTATUS")
        if self._is_blank(raw):
            return True, ""
        if str(raw).strip() in {"2", "02"}:
            return True, ""
        return False, self.XPLANT_REASON

    def get_rules(self) -> dict:
        return {
            "MATERIALNUMBER":     self.validate_material_number,
            "PLANT":              self.validate_plant,
            "PRODUCTDESCRIPTION": self.validate_product_description,
            "PRODUCTTYPE":        self.validate_product_type,
            "PRODUCTHIERARCHY":   self.validate_product_hierarchy,
            "MRPTYPE":            self.validate_mrp_type,
            "PROCUREMENTTYPE":    self.validate_procurement_type,
            "ABCINDICATOR":       self.validate_abc_indicator,
            "IBPSTATUS":          self.validate_ibp_status,
            "XPLANTMATSTATUS":    self.validate_xplant_mat_status,
        }


# ══════════════════════════════════════════════
#  Validator
# ══════════════════════════════════════════════
class PartTableValidator:

    def __init__(self, filepath: str, valid_plants: set):
        self.filepath     = filepath
        self.valid_plants = valid_plants
        self.df           = pd.DataFrame()
        self.error_map    = {}

    def load(self):
        if self.filepath.lower().endswith(".csv"):
            self.df = pd.read_csv(self.filepath, dtype=str)
        else:
            self.df = pd.read_excel(self.filepath, dtype=str)
        self.df.columns = [c.strip().upper() for c in self.df.columns]

    def validate(self):
        engine = RuleEngine(self.valid_plants)
        rules  = engine.get_rules()

        for idx, row in self.df.iterrows():
            errors = {}
            for col, rule_fn in rules.items():
                try:
                    passed, reason = rule_fn(row)
                except Exception as e:
                    passed, reason = False, f"Exception: {str(e)}"
                if not passed:
                    errors[col] = reason
            if errors:
                self.error_map[idx] = errors


# ══════════════════════════════════════════════
#  Report Writer
# ══════════════════════════════════════════════
class ReportWriter:

    SHEET_SUMMARY = "Summary"
    SHEET_RULES   = "Rules"

    RULES_CONTENT = {
        "MATERIALNUMBER": [
            "Must not be blank.",
            "For FERT type: Material number must be in range 14000000000000 – 14999999999999.",
            "For HAWA type: Material number must be in range 15000000000000 – 15999999999999.",
        ],
        "PLANT": [
            "Must not be blank.",
            "Must be present in the Consolidated PL list (loaded from SharePoint).",
        ],
        "PRODUCTDESCRIPTION": ["Must not be blank."],
        "PRODUCTTYPE": [
            "Must not be blank.",
            "Value must be either FERT or HAWA.",
        ],
        "PRODUCTHIERARCHY": ["Must not be blank."],
        "MRPTYPE": [
            "Must not be blank.",
            "Value must be either ND or PD.",
        ],
        "PROCUREMENTTYPE": ["Must not be blank."],
        "ABCINDICATOR": [
            "Must not be blank.",
            "NOTE: Column is fully blank in source data – defaulting to 100% error.",
        ],
        "IBPSTATUS": [
            "Allowed values: IBP or blank.",
            "Any other value is treated as an error.",
        ],
        "XPLANTMATSTATUS": [
            "Allowed values: 2 or blank.",
            "Any other value is treated as an error.",
        ],
    }

    def __init__(self, validator: PartTableValidator, output_path: str):
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
            ws.column_dimensions[get_column_letter(col[0].column)].width = min(max(length + 3, min_w), max_w)

    # ── Summary sheet ────────────────────────
    def _write_summary_sheet(self, wb, error_map: dict, total_rows: int):
        ws = wb.create_sheet(self.SHEET_SUMMARY)

        col_error_counts = {col: 0 for col in RULES_FIELDS_ORDERED}
        for bad_cols in error_map.values():
            for col in bad_cols.keys():
                if col in col_error_counts:
                    col_error_counts[col] += 1

        REASON_MAP = {
            "MATERIALNUMBER":     "MATERIALNUMBER: Must be 14xxxxxxxxxxxxxx (FERT) or 15xxxxxxxxxxxxxx (HAWA) — invalid range or blank",
            "PLANT":              "PLANT: Value is not in the Consolidated PL list",
            "PRODUCTDESCRIPTION": "PRODUCTDESCRIPTION: Field is blank",
            "PRODUCTTYPE":        "PRODUCTTYPE: Must be FERT or HAWA — invalid or blank value found",
            "PRODUCTHIERARCHY":   "PRODUCTHIERARCHY: Field is blank",
            "MRPTYPE":            "MRPTYPE: Field is blank",
            "PROCUREMENTTYPE":    "PROCUREMENTTYPE: Field is blank",
            "ABCINDICATOR":       "ABCINDICATOR: Field is blank",
            "IBPSTATUS":          "IBPSTATUS: Must be 'IBP' or blank — unexpected value found",
            "XPLANTMATSTATUS":    "XPLANTMATSTATUS: Must be '2' or blank — unexpected value found",
        }

        # Title
        ws.merge_cells("A1:G1")
        title_cell           = ws.cell(row=1, column=1, value="Part Master FG Validation Summary")
        title_cell.font      = Font(name="Arial", bold=True, size=14)
        title_cell.fill      = TITLE_FILL
        title_cell.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[1].height = 24

        headers = ["#", "Field Name", "Error Count", "Record Count", "% Health", "% of Error", "Reason"]
        for c_idx, h in enumerate(headers, start=1):
            cell           = ws.cell(row=2, column=c_idx, value=h)
            cell.fill      = TITLE_FILL
            cell.font      = Font(name="Arial", bold=True)
            cell.border    = THIN_BORDER
            cell.alignment = Alignment(horizontal="center", vertical="center")

        row_num = 3
        for field_num, col_name in enumerate(RULES_FIELDS_ORDERED, start=1):
            count      = col_error_counts.get(col_name, 0)
            pct_error  = round((count / total_rows) * 100, 2) if total_rows else 0
            pct_health = round(100 - pct_error, 2)
            reason_text = REASON_MAP.get(col_name, "") if count > 0 else ""

            ws.cell(row=row_num, column=1, value=field_num)
            ws.cell(row=row_num, column=2, value=col_name)
            ws.cell(row=row_num, column=3, value=count)
            ws.cell(row=row_num, column=4, value=total_rows)
            ws.cell(row=row_num, column=5, value=f"{pct_health}%")
            ws.cell(row=row_num, column=6, value=f"{pct_error}%")
            ws.cell(row=row_num, column=7, value=reason_text)

            for c in range(1, 8):
                ws.cell(row=row_num, column=c).font      = BODY_FONT
                ws.cell(row=row_num, column=c).border    = THIN_BORDER
                align = Alignment(horizontal="left" if c == 7 else "center", vertical="center")
                ws.cell(row=row_num, column=c).alignment = align
            row_num += 1

        # TOTAL row
        total_errors       = sum(col_error_counts.values())
        total_record_count = total_rows * len(RULES_FIELDS_ORDERED)
        total_pct_error    = round((total_errors / total_record_count) * 100, 2) if total_record_count else 0
        total_pct_health   = round(100 - total_pct_error, 2)

        for c_idx, val in enumerate(
            ["", "TOTAL", total_errors, total_record_count, f"{total_pct_health}%", f"{total_pct_error}%", ""],
            start=1,
        ):
            cell           = ws.cell(row=row_num, column=c_idx, value=val)
            cell.font      = Font(name="Arial", bold=True)
            cell.fill      = TOTAL_FILL
            cell.border    = THIN_BORDER
            cell.alignment = Alignment(horizontal="center", vertical="center")

        row_num += 2

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
            value_cell.font      = BODY_FONT
            value_cell.border    = THIN_BORDER
            value_cell.alignment = Alignment(horizontal="center", vertical="center")
            row_num += 1

        self._auto_width(ws, min_w=8, max_w=70)

    # ── Rule_Set sheet ────────────────────────
    def _write_ruleset_sheet(self, wb):
        ws = wb.create_sheet(self.SHEET_RULES)

        ws.merge_cells("A1:C1")
        title_cell           = ws.cell(row=1, column=1, value="Part Master FG – Validation Rules")
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

        for field, rules_list in self.RULES_CONTENT.items():
            num_rules = len(rules_list)
            for r_idx, rule_text in enumerate(rules_list):
                ws.cell(row=current_row, column=1,
                        value=rule_num if r_idx == 0 else "").fill = RULE_FILL
                ws.cell(row=current_row, column=1).font      = Font(name="Arial", size=10, bold=(r_idx == 0))
                ws.cell(row=current_row, column=1).border    = THIN_BORDER
                ws.cell(row=current_row, column=1).alignment = Alignment(horizontal="center", vertical="center")

                ws.cell(row=current_row, column=2,
                        value=field if r_idx == 0 else "").fill = RULE_FILL
                ws.cell(row=current_row, column=2).font      = Font(name="Arial", size=10, bold=(r_idx == 0))
                ws.cell(row=current_row, column=2).border    = THIN_BORDER
                ws.cell(row=current_row, column=2).alignment = Alignment(horizontal="center", vertical="center")

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

    # ── Field Error Sheets ───────────────────
    def _write_field_error_sheets(self, wb, df: pd.DataFrame):
        """
        Creates one error sub-sheet per failing field.

        Changes vs original:
        1. Only columns defined in RULES_CONTENT (+ ERROR_COLUMNS) are shown.
        2. Sheet creation order: ERROR_SHEET_PRIORITY first, then remaining fields.
        3. ABCINDICATOR sheet still skipped (force-fail rule, not a per-row sheet).
        """
        v = self.validator

        # Columns that may appear in error sheets = only those in rules sheet
        rules_cols_in_data = [c for c in RULES_FIELDS_ORDERED if c in df.columns]

        all_error_fields = set()
        for bad_cols in v.error_map.values():
            all_error_fields.update(bad_cols.keys())

        # Build the sheet creation order:
        # priority fields first (only if they have errors), then remaining alphabetically
        ordered_fields = []
        for f in ERROR_SHEET_PRIORITY:
            if f in all_error_fields and f != "ABCINDICATOR":
                ordered_fields.append(f)
        for f in sorted(all_error_fields):
            if f not in ordered_fields and f != "ABCINDICATOR":
                ordered_fields.append(f)

        for field_name in ordered_fields:
            row_indices = [idx for idx, errdict in v.error_map.items() if field_name in errdict]
            if not row_indices:
                continue

            # ── Only keep columns that appear in the Rules sheet ──
            display_cols = [c for c in rules_cols_in_data]   # ordered subset
            subset = df.loc[row_indices, display_cols].copy()

            subset["ERROR_COLUMNS"] = subset.index.map(
                lambda i: v.error_map.get(i, {}).get(field_name, "")
            )

            # Sheet name (Excel limit: 31 chars)
            sheet_name = field_name[:31]
            existing   = [s.title for s in wb.worksheets]
            counter    = 1
            base_name  = sheet_name
            while sheet_name in existing:
                sheet_name = f"{base_name[:28]}_{counter}"
                counter   += 1

            ws = wb.create_sheet(sheet_name)
            self._write_header(ws, subset.columns)

            col_idx_map = {col: i for i, col in enumerate(subset.columns, start=1)}

            for r_idx, (orig_idx, row_data) in enumerate(subset.iterrows(), start=2):
                for c_idx, (col, value) in enumerate(zip(subset.columns, row_data), start=1):
                    cell           = ws.cell(row=r_idx, column=c_idx, value=value)
                    cell.font      = BODY_FONT
                    cell.border    = THIN_BORDER
                    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                    cell.fill      = ROW_FILL

                # Highlight the erroring column in red
                if field_name in col_idx_map:
                    target_cell      = ws.cell(row=r_idx, column=col_idx_map[field_name])
                    target_cell.fill = RED_FILL
                    target_cell.font = ERR_FONT

            self._auto_width(ws, min_w=10, max_w=60)
            ws.freeze_panes = "A2"

            note_row = len(subset) + 3
            ws.cell(
                row=note_row, column=1,
                value=f"Total error rows for '{field_name}': {len(subset)}",
            ).font = Font(name="Arial", italic=True, size=9, bold=True)

    # ── Orchestrate Writing ──────────────────
    def write(self):
        v  = self.validator
        df = v.df.copy()

        wb = Workbook()
        if "Sheet" in wb.sheetnames:
            del wb["Sheet"]

        self._write_summary_sheet(wb, v.error_map, total_rows=len(df))
        self._write_ruleset_sheet(wb)
        self._write_field_error_sheets(wb, df)

        wb.save(self.output_path)
        print(f"\n✅  Output saved  → {self.output_path}")
        print(f"   Total rows    : {len(df)}")
        print(f"   Error rows    : {len(v.error_map)}")


# ══════════════════════════════════════════════
#  Orchestrator
# ══════════════════════════════════════════════
class PartTableProcessor:

    def __init__(self, input_path: str, output_path: str):
        # Load plant codes live from SharePoint
        valid_plants = load_valid_plants_from_sharepoint(
            sharepoint_url  = SHAREPOINT_URL,
            sharepoint_site = SHAREPOINT_SITE,
            file_path       = SHAREPOINT_FILE_PATH,
            email           = SHAREPOINT_EMAIL,
            password        = SHAREPOINT_PASSWORD,
            sheet_name      = SHAREPOINT_SHEET,
            column_name     = SHAREPOINT_COL,
        )
        self.validator = PartTableValidator(input_path, valid_plants)
        self.writer    = ReportWriter(self.validator, output_path)

    def run(self):
        print("📂  Loading file …")
        self.validator.load()
        print(f"    Columns detected : {list(self.validator.df.columns)}")
        print("🔍  Validating rules …")
        self.validator.validate()
        print("📝  Writing report …")
        self.writer.write()


# ══════════════════════════════════════════════
#  ENTRY POINT
# ══════════════════════════════════════════════
if __name__ == "__main__":
    processor = PartTableProcessor(INPUT_FILE, OUTPUT_FILE)
    processor.run()
