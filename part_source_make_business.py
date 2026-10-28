import re
from collections import defaultdict
from datetime import datetime
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ─────────────────────────────────────────────
#  FILE PATHS
# ─────────────────────────────────────────────
PARTSOURCE_MAKE_INPUT_FILE   = r"C:\Users\SW526XH\Downloads\Go Live-2\PartSource(Make)\PartSource_Make.tab"
BOM_INPUT_FILE                = r"C:\Users\SW526XH\Downloads\Go Live-2\BOM\BOM.tab"
SOURCE_CONSTRAINT_INPUT_FILE  = r"C:\Users\SW526XH\Downloads\Go Live-2\SourceConstraint\SourceConstraint.tab"
PARTMASTER_INPUT_FILE         = r"C:\Users\SW526XH\Downloads\Go Live-2\Part\Part.tab"   # <-- update to your Part master file
OUTPUT_FILE                   = r"C:\Users\SW526XH\Downloads\Go Live-2\PartSource(Make)\Validated_PartSourceMake_Business.xlsx"


# ─────────────────────────────────────────────
#  PARTSOURCE(MAKE) COLUMN NAMES  (after .strip().upper())
# ─────────────────────────────────────────────
PS_MATERIAL_COL = "MATERIAL"
PS_PLANT_COL    = "PLANT"

# ─────────────────────────────────────────────
#  BOM MASTER COLUMN NAMES  (after .strip().upper())
# ─────────────────────────────────────────────
BOM_MATERIAL_COL = "ROOT_MATERIAL"
BOM_PLANT_COL    = "ROOT_PLANT"

# ─────────────────────────────────────────────
#  SOURCE CONSTRAINT MASTER COLUMN NAMES  (after .strip().upper())
# ─────────────────────────────────────────────
SC_MATERIAL_COL = "MATERIAL"
SC_PLANT_COL    = "MKALPLANT"
SC_FACTOR_COL   = "CONSTRAINTFACTOR"

# ─────────────────────────────────────────────
#  PART MASTER COLUMN NAMES  (after .strip().upper())
#  Adjust here if the Part master uses different names
# ─────────────────────────────────────────────
PM_MATERIAL_COL = "MATERIALNUMBER"
PM_PLANT_COL    = "PLANT"
# Whichever of these exists in the Part file is used (handles the PROCUREMENTTYPE / PROCUREMENTYPE spelling)
PM_PROCTYPE_CANDIDATES = ["PROCUREMENTTYPE", "PROCUREMENTYPE"]

# Required PROCUREMENTTYPE in Part master for every PartSource(Make) Part-Site
REQUIRED_PROCUREMENT_TYPE = "E"

# Sub-rule (a) behaviour when the Part-Site is NOT present in the Source Constraint extract at all:
#   False -> not flagged again here (already reported by the "Not in Source Constraint" sub-category)
#   True  -> also flagged as "CONSTRAINTFACTOR not maintained"
FACTOR_MISSING_COMBO_IS_ERROR = False


# ─────────────────────────────────────────────
#  Colours / Styles  — matched to existing Business Validator template
# ─────────────────────────────────────────────
RED_FILL        = PatternFill("solid", start_color="FF0000", end_color="FF0000")
HDR_FILL        = PatternFill("solid", start_color="D9E1F2", end_color="D9E1F2")
RULE_FILL       = PatternFill("solid", start_color="E2EFDA", end_color="E2EFDA")
TITLE_FILL      = PatternFill("solid", start_color="BDD7EE", end_color="BDD7EE")
TOTAL_FILL      = PatternFill("solid", start_color="F2F2F2", end_color="F2F2F2")
WHITE_FILL      = PatternFill("solid", start_color="FFFFFF", end_color="FFFFFF")
STATS_FILL      = PatternFill("solid", start_color="EDEDED", end_color="EDEDED")
SUMM_HDR_FILL   = PatternFill("solid", start_color="BDD7EE", end_color="BDD7EE")
SUMM_TITLE_FILL = PatternFill("solid", start_color="BDD7EE", end_color="BDD7EE")
SUBROW_FILL     = PatternFill("solid", start_color="F7F7F7", end_color="F7F7F7")

HDR_FONT     = Font(bold=True, name="Arial")
BODY_FONT    = Font(name="Arial", size=10)
ERR_FONT     = Font(name="Arial", size=10, bold=True, color="FFFFFF")
SUBROW_FONT  = Font(name="Arial", size=10, italic=True)
THIN_BORDER  = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"),  bottom=Side(style="thin"),
)

# ─────────────────────────────────────────────
#  FIELD ORDER  (only fields with business rules)
#  ACTIVE: MATERIAL only.  All other rules are commented out below.
# ─────────────────────────────────────────────
# FIELD_ORDER = ["MAXIMUMLOTSIZE", "MINIMUMLOTSIZE", "ROUNDINGVALUE", "MATERIAL",
#                "MATERIAL_PROCUREMENTTYPE"]
FIELD_ORDER = ["MATERIAL"]

# Single-line reason shown in Summary sheet for the simple fields
FIELD_REASON = {
    # "MAXIMUMLOTSIZE": "MAXIMUMLOTSIZE: Value not greater than 0",
    # "MINIMUMLOTSIZE": "MINIMUMLOTSIZE: Value not greater than 0",
    # "ROUNDINGVALUE":  "ROUNDINGVALUE: Value not greater than 0",
    # "MATERIAL_PROCUREMENTTYPE": (
    #     f"MATERIAL_PROCUREMENTTYPE: Part-Site PROCUREMENTTYPE is not "
    #     f"'{REQUIRED_PROCUREMENT_TYPE}' in Part master"
    # ),
}

# MATERIAL sub-category keys (used as counters + Summary sub-rows)
SUB_BLANK  = "Blank"
SUB_NO_BOM = "Not in BOM"
SUB_NO_SC  = "Not in Source Constraint"
SUB_FACTOR = "CONSTRAINTFACTOR not maintained / Zero"
SUB_PROC   = f"PROCUREMENTTYPE not '{REQUIRED_PROCUREMENT_TYPE}' in Part"

MATERIAL_SUB_CATEGORIES = [SUB_BLANK, SUB_NO_BOM, SUB_NO_SC, SUB_FACTOR, SUB_PROC]


# ══════════════════════════════════════════════
#  Rule Engine
# ══════════════════════════════════════════════
class PartSourceMakeBusinessRuleEngine:
    """
    ACTIVE RULE:
      MATERIAL
        - The MATERIAL / PLANT (Part-Site) combination must be present in the
          BOM extract as ROOT_MATERIAL / ROOT_PLANT, AND must also be present
          in the Source Constraint extract as MATERIAL / MKALPLANT.
        a) The Part-Site combination must have CONSTRAINTFACTOR maintained
           (not blank) and it must not be zero in the Source Constraint extract.
        b) The Part-Site combination must have PROCUREMENTTYPE maintained as
           "E" in the Part extract.

    COMMENTED OUT (kept for reference):
      MAXIMUMLOTSIZE / MINIMUMLOTSIZE / ROUNDINGVALUE - must be > 0.
      MATERIAL_PROCUREMENTTYPE - superseded by MATERIAL sub-rule (b).
    """

    def __init__(self, valid_bom_combos: set, valid_sc_combos: set,
                 partmaster_lookup=None, sc_factor_lookup=None):
        self.valid_bom_combos  = valid_bom_combos
        self.valid_sc_combos   = valid_sc_combos
        # {(material, plant): {procurement types found in Part master}}
        # None -> sub-rule (b) is skipped
        self.partmaster_lookup = partmaster_lookup
        # {(material, plant): [CONSTRAINTFACTOR values found in Source Constraint]}
        # None -> sub-rule (a) is skipped
        self.sc_factor_lookup  = sc_factor_lookup

    @staticmethod
    def _is_blank(value) -> bool:
        return pd.isna(value) or str(value).strip() == ""

    @staticmethod
    def _parse_number(value):
        s = str(value).strip()
        if s == "" or s.lower() == "nan":
            return None
        try:
            return float(s)
        except ValueError:
            return None

    # ── COMMENTED OUT: numeric > 0 rules ──────────────────────────────
    # def _validate_positive_field(self, row, field_name: str) -> str:
    #     """Returns a reason string if field_name is present but not > 0, else ''."""
    #     raw = row.get(field_name, "")
    #     num = self._parse_number(raw)
    #     if num is not None and num <= 0:
    #         return f"{field_name}: '{str(raw).strip()}' is not greater than 0"
    #     return ""

    def _validate_material(self, row):
        """Returns (reason_string, [sub_category_keys]).
        reason_string is '' (and the list empty) if MATERIAL passes."""
        mat_raw   = row.get(PS_MATERIAL_COL, "")
        plant_raw = row.get(PS_PLANT_COL, "")
        mat_blank   = self._is_blank(mat_raw)
        plant_blank = self._is_blank(plant_raw)

        if mat_blank or plant_blank:
            missing = []
            if mat_blank:
                missing.append(PS_MATERIAL_COL)
            if plant_blank:
                missing.append(PS_PLANT_COL)
            return f"MATERIAL: Blank field(s) - {', '.join(missing)}", [SUB_BLANK]

        mat   = str(mat_raw).strip()
        plant = str(plant_raw).strip()
        key   = f"{mat}|{plant}"

        problems = []
        subs     = []

        # ── Existing: presence in BOM ──
        if key not in self.valid_bom_combos:
            problems.append("not found in BOM extract (ROOT_MATERIAL/ROOT_PLANT)")
            subs.append(SUB_NO_BOM)

        # ── Existing: presence in Source Constraint ──
        if key not in self.valid_sc_combos:
            problems.append("not found in Source Constraint extract (MATERIAL/MKALPLANT)")
            subs.append(SUB_NO_SC)

        # ── NEW (a): CONSTRAINTFACTOR maintained and not zero ──
        if self.sc_factor_lookup is not None:
            factors = self.sc_factor_lookup.get((mat, plant))
            if factors is None:
                if FACTOR_MISSING_COMBO_IS_ERROR:
                    problems.append(
                        f"{SC_FACTOR_COL} not maintained (combination not present in "
                        f"Source Constraint extract)"
                    )
                    subs.append(SUB_FACTOR)
            else:
                has_blank = any(f == "" for f in factors)
                has_zero  = any((self._parse_number(f) == 0) for f in factors if f != "")
                if has_blank and has_zero:
                    problems.append(f"{SC_FACTOR_COL} is blank and also '0' in Source Constraint extract")
                    subs.append(SUB_FACTOR)
                elif has_blank:
                    problems.append(f"{SC_FACTOR_COL} is not maintained (blank) in Source Constraint extract")
                    subs.append(SUB_FACTOR)
                elif has_zero:
                    problems.append(f"{SC_FACTOR_COL} is '0' in Source Constraint extract")
                    subs.append(SUB_FACTOR)

        # ── NEW (b): PROCUREMENTTYPE must be 'E' in Part extract ──
        if self.partmaster_lookup is not None:
            found = self.partmaster_lookup.get((mat, plant))
            if found is None:
                problems.append(
                    f"not found in Part extract (expected PROCUREMENTTYPE "
                    f"'{REQUIRED_PROCUREMENT_TYPE}')"
                )
                subs.append(SUB_PROC)
            elif found != {REQUIRED_PROCUREMENT_TYPE}:
                actual = ", ".join(sorted(v if v else "blank" for v in found))
                problems.append(
                    f"PROCUREMENTTYPE is '{actual}' in Part extract "
                    f"(expected '{REQUIRED_PROCUREMENT_TYPE}')"
                )
                subs.append(SUB_PROC)

        if problems:
            return f"MATERIAL: Part-Site combination '{mat}-{plant}' " + "; ".join(problems), subs

        return "", []

    # ── COMMENTED OUT: standalone MATERIAL_PROCUREMENTTYPE rule
    #    (logic now lives inside MATERIAL sub-rule (b) above) ──────────
    # def _validate_material_procurement_type(self, row) -> str:
    #     if self.partmaster_lookup is None:
    #         return ""
    #     mat_raw   = row.get(PS_MATERIAL_COL, "")
    #     plant_raw = row.get(PS_PLANT_COL, "")
    #     if self._is_blank(mat_raw) or self._is_blank(plant_raw):
    #         return ""
    #     mat   = str(mat_raw).strip()
    #     plant = str(plant_raw).strip()
    #     found = self.partmaster_lookup.get((mat, plant))
    #     if found is None:
    #         return (
    #             f"MATERIAL_PROCUREMENTTYPE: Part-Site '{mat}-{plant}' not found in "
    #             f"Part master (expected PROCUREMENTTYPE '{REQUIRED_PROCUREMENT_TYPE}')"
    #         )
    #     if found != {REQUIRED_PROCUREMENT_TYPE}:
    #         actual = ", ".join(sorted(v if v else "blank" for v in found))
    #         return (
    #             f"MATERIAL_PROCUREMENTTYPE: Part-Site '{mat}-{plant}' has "
    #             f"PROCUREMENTTYPE '{actual}' in Part master "
    #             f"(expected '{REQUIRED_PROCUREMENT_TYPE}')"
    #         )
    #     return ""

    def validate_row(self, row):
        """Returns ({field_name: reason}, [material_sub_category_keys])."""
        reasons = {}

        # ── COMMENTED OUT: numeric > 0 rules ──
        # for field_name in ("MAXIMUMLOTSIZE", "MINIMUMLOTSIZE", "ROUNDINGVALUE"):
        #     reason = self._validate_positive_field(row, field_name)
        #     if reason:
        #         reasons[field_name] = reason

        material_reason, material_subs = self._validate_material(row)
        if material_reason:
            reasons["MATERIAL"] = material_reason

        # ── COMMENTED OUT: standalone MATERIAL_PROCUREMENTTYPE rule ──
        # proc_reason = self._validate_material_procurement_type(row)
        # if proc_reason:
        #     reasons["MATERIAL_PROCUREMENTTYPE"] = proc_reason

        return reasons, material_subs


# ══════════════════════════════════════════════
#  Validator
# ══════════════════════════════════════════════
class PartSourceMakeBusinessTableValidator:

    def __init__(self, ps_make_path: str, bom_path: str, source_constraint_path: str,
                 partmaster_path: str):
        self.ps_make_path           = ps_make_path
        self.bom_path               = bom_path
        self.source_constraint_path = source_constraint_path
        self.partmaster_path        = partmaster_path

        self.df                = pd.DataFrame()
        self.pm_df             = pd.DataFrame()
        self.valid_bom_combos  = set()
        self.valid_sc_combos   = set()
        self.sc_factor_lookup  = None
        self.partmaster_lookup = None
        self.pm_proctype_col   = None
        self.error_map         = {}   # row_idx -> [failed field names]
        self.reason_map        = {}   # row_idx -> {field: reason}

        # MATERIAL sub-category counters for Summary sheet breakdown
        self.material_sub_counts = {k: 0 for k in MATERIAL_SUB_CATEGORIES}

    @staticmethod
    def _clean(series: pd.Series) -> pd.Series:
        return series.fillna("").astype(str).str.strip()

    def load(self):
        self.df = pd.read_csv(self.ps_make_path, sep="\t", dtype=str)
        self.df.columns = [c.strip().upper() for c in self.df.columns]

        # ── BOM ──
        bom_df = pd.read_csv(self.bom_path, sep="\t", dtype=str)
        bom_df.columns = [c.strip().upper() for c in bom_df.columns]
        bom_mat   = self._clean(bom_df[BOM_MATERIAL_COL])
        bom_plant = self._clean(bom_df[BOM_PLANT_COL])
        for mat, plant in zip(bom_mat, bom_plant):
            if mat == "" or plant == "":
                continue
            self.valid_bom_combos.add(f"{mat}|{plant}")

        # ── Source Constraint (combos + CONSTRAINTFACTOR values) ──
        sc_df = pd.read_csv(self.source_constraint_path, sep="\t", dtype=str)
        sc_df.columns = [c.strip().upper() for c in sc_df.columns]
        sc_mat   = self._clean(sc_df[SC_MATERIAL_COL])
        sc_plant = self._clean(sc_df[SC_PLANT_COL])
        for mat, plant in zip(sc_mat, sc_plant):
            if mat == "" or plant == "":
                continue
            self.valid_sc_combos.add(f"{mat}|{plant}")

        self._build_sc_factor_lookup(sc_df, sc_mat, sc_plant)

        # ── Part master ──
        self.pm_df = pd.read_csv(self.partmaster_path, sep="\t", dtype=str)
        self.pm_df.columns = [c.strip().upper() for c in self.pm_df.columns]
        self._build_partmaster_lookup()

    def _build_sc_factor_lookup(self, sc_df, sc_mat, sc_plant):
        """Builds {(material, plant): [CONSTRAINTFACTOR values]} from Source Constraint.
        Leaves the lookup as None (sub-rule (a) skipped) if CONSTRAINTFACTOR is missing."""
        if SC_FACTOR_COL not in sc_df.columns:
            print(f"⚠️   MATERIAL sub-rule (a) skipped — Source Constraint column not found: {SC_FACTOR_COL}")
            self.sc_factor_lookup = None
            return

        sc_factor = self._clean(sc_df[SC_FACTOR_COL])
        lookup = defaultdict(list)
        for m, p, f in zip(sc_mat, sc_plant, sc_factor):
            if m and p:
                lookup[(m, p)].append(f)
        self.sc_factor_lookup = dict(lookup)

    def _build_partmaster_lookup(self):
        """Builds {(material, plant): {procurement types}} from the Part master.
        Leaves the lookup as None (sub-rule (b) skipped) if required columns are missing."""
        self.pm_proctype_col = next(
            (c for c in PM_PROCTYPE_CANDIDATES if c in self.pm_df.columns), None
        )

        ps_missing = [c for c in (PS_MATERIAL_COL, PS_PLANT_COL)
                      if c not in self.df.columns]
        pm_missing = [c for c in (PM_MATERIAL_COL, PM_PLANT_COL)
                      if c not in self.pm_df.columns]
        if self.pm_proctype_col is None:
            pm_missing.append("/".join(PM_PROCTYPE_CANDIDATES))

        if ps_missing or pm_missing:
            if ps_missing:
                print(f"⚠️   MATERIAL sub-rule (b) skipped — PartSource(Make) column(s) not found: {ps_missing}")
            if pm_missing:
                print(f"⚠️   MATERIAL sub-rule (b) skipped — Part master column(s) not found: {pm_missing}")
            self.partmaster_lookup = None
            return

        mats   = self._clean(self.pm_df[PM_MATERIAL_COL])
        plants = self._clean(self.pm_df[PM_PLANT_COL])
        ptypes = self._clean(self.pm_df[self.pm_proctype_col])

        lookup = defaultdict(set)
        for m, p, t in zip(mats, plants, ptypes):
            if m and p:
                lookup[(m, p)].add(t)

        self.partmaster_lookup = dict(lookup)

    def validate(self):
        engine = PartSourceMakeBusinessRuleEngine(
            self.valid_bom_combos, self.valid_sc_combos,
            self.partmaster_lookup, self.sc_factor_lookup
        )

        for idx, row in self.df.iterrows():
            try:
                reasons, material_subs = engine.validate_row(row)
            except Exception:
                reasons, material_subs = {}, []

            if reasons:
                self.error_map[idx]  = list(reasons.keys())
                self.reason_map[idx] = reasons

                for sub in material_subs:
                    self.material_sub_counts[sub] += 1

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


# ══════════════════════════════════════════════
#  Report Writer
# ══════════════════════════════════════════════
class PartSourceMakeBusinessReportWriter:

    SHEET_SUMMARY = "Summary"
    SHEET_RULES   = "Rules"

    RULES_CONTENT = {
        # "MAXIMUMLOTSIZE": [
        #     "MAXIMUMLOTSIZE must be greater than 0.",
        # ],
        # "MINIMUMLOTSIZE": [
        #     "MINIMUMLOTSIZE must be greater than 0.",
        # ],
        # "ROUNDINGVALUE": [
        #     "ROUNDINGVALUE must be greater than 0.",
        # ],
        "MATERIAL": [
            "The MATERIAL / PLANT (Part-Site) combination must be present in the "
            "BOM extract as ROOT_MATERIAL / ROOT_PLANT.",
            "The MATERIAL / PLANT (Part-Site) combination must also be present in "
            "the Source Constraint extract as MATERIAL / MKALPLANT.",
            "a) The Part-Site combination here should have CONSTRAINTFACTOR maintained "
            "in the Source Constraint extract and it should not be zero.",
            "b) The Part-Site combination here should have PROCUREMENTTYPE maintained "
            f"as \"{REQUIRED_PROCUREMENT_TYPE}\" in the Part extract.",
        ],
        # "MATERIAL_PROCUREMENTTYPE": [
        #     "For a Part-Site combination here (MATERIAL and PLANT columns in the "
        #     "PartSource(Make) extract), the part-site's PROCUREMENTTYPE should be "
        #     f"maintained as \"{REQUIRED_PROCUREMENT_TYPE}\" in the Part master.",
        # ],
    }

    def __init__(self, validator: PartSourceMakeBusinessTableValidator, output_path: str):
        self.validator   = validator
        self.output_path = output_path

    # ── helpers ──────────────────────────────
    def _write_header(self, ws, columns):
        for c_idx, col_name in enumerate(columns, start=1):
            cell = ws.cell(row=1, column=c_idx, value=col_name)
            if col_name == "ERROR_COLUMNS":
                cell.fill = WHITE_FILL
                cell.font = Font(bold=True, name="Arial", color="000000")
            else:
                cell.fill = HDR_FILL
                cell.font = HDR_FONT
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border    = THIN_BORDER

    def _set_widths(self, ws):
        for col in ws.columns:
            max_len = max((len(str(c.value)) if c.value else 0) for c in col)
            ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 4, 60)

    def _style_summary_data_row(self, ws, row_num: int, num_cols: int = 7,
                                bold: bool = False, fill: PatternFill = None,
                                italic: bool = False, indent_col2: bool = False):
        for c in range(1, num_cols + 1):
            cell           = ws.cell(row=row_num, column=c)
            cell.font      = Font(name="Arial", bold=bold, italic=italic, size=10)
            cell.border    = THIN_BORDER
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if fill:
                cell.fill = fill
        if indent_col2:
            ws.cell(row=row_num, column=2).alignment = Alignment(
                horizontal="left", vertical="center", indent=2
            )

    # ══════════════════════════════════════════
    #  Summary sheet
    # ══════════════════════════════════════════
    def _write_summary_sheet_into(self, ws, error_map: dict, total_rows: int):
        v = self.validator

        # ── Row 1 : Title ──
        ws.merge_cells("A1:G1")
        title_cell           = ws.cell(row=1, column=1,
                                        value="PartSource(Make) Business Validation Summary")
        title_cell.font      = Font(name="Arial", bold=True, size=14)
        title_cell.fill      = SUMM_TITLE_FILL
        title_cell.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[1].height = 26

        # ── Row 2 : Column headers ──
        headers = ["#", "Field Name", "Error Count", "Record Count",
                   "% Health", "% of Error", "Reason / Sub-Category"]
        for c_idx, h in enumerate(headers, start=1):
            cell           = ws.cell(row=2, column=c_idx, value=h)
            cell.fill      = SUMM_HDR_FILL
            cell.font      = Font(name="Arial", bold=True, size=10)
            cell.border    = THIN_BORDER
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.row_dimensions[2].height = 30

        # ── Per-field error counts ──
        col_error_counts: dict = {}
        for bad_cols in error_map.values():
            for col in bad_cols:
                col_error_counts[col] = col_error_counts.get(col, 0) + 1

        row_num   = 3
        field_num = 1

        for col_name in FIELD_ORDER:
            count      = col_error_counts.get(col_name, 0)
            has_errors = count > 0

            pct_error  = round((count / total_rows) * 100, 2) if total_rows else 0
            pct_health = round(100 - pct_error, 2)

            if col_name == "MATERIAL":
                reason_text = ("Blank / not found in BOM (ROOT_MATERIAL-ROOT_PLANT) / "
                               "not found in Source Constraint (MATERIAL-MKALPLANT) / "
                               f"CONSTRAINTFACTOR not maintained or zero / "
                               f"PROCUREMENTTYPE not '{REQUIRED_PROCUREMENT_TYPE}' in Part") if has_errors else ""
            else:
                reason_text = FIELD_REASON.get(col_name, "") if has_errors else ""

            ws.cell(row=row_num, column=1, value=field_num)
            ws.cell(row=row_num, column=2, value=col_name)
            ws.cell(row=row_num, column=3, value=count)
            ws.cell(row=row_num, column=4, value=total_rows)
            ws.cell(row=row_num, column=5, value=f"{pct_health}%")
            ws.cell(row=row_num, column=6, value=f"{pct_error}%")
            ws.cell(row=row_num, column=7, value=reason_text)

            self._style_summary_data_row(ws, row_num, bold=(col_name == "MATERIAL"), fill=WHITE_FILL)
            ws.cell(row=row_num, column=7).alignment = Alignment(
                horizontal="left", vertical="center", wrap_text=True
            )
            row_num += 1
            field_num += 1

            # ── MATERIAL sub-rows ──
            if col_name == "MATERIAL":
                for sub_label in MATERIAL_SUB_CATEGORIES:
                    sub_count      = v.material_sub_counts.get(sub_label, 0)
                    sub_pct_error  = round((sub_count / total_rows) * 100, 2) if total_rows else 0
                    sub_pct_health = round(100 - sub_pct_error, 2)

                    ws.cell(row=row_num, column=1, value="")
                    ws.cell(row=row_num, column=2, value=f"└─ {sub_label}")
                    ws.cell(row=row_num, column=3, value=sub_count)
                    ws.cell(row=row_num, column=4, value=total_rows)
                    ws.cell(row=row_num, column=5, value=f"{sub_pct_health}%")
                    ws.cell(row=row_num, column=6, value=f"{sub_pct_error}%")
                    ws.cell(row=row_num, column=7, value="")
                    self._style_summary_data_row(ws, row_num, italic=True, fill=SUBROW_FILL,
                                                  indent_col2=True)
                    row_num += 1

        # ── TOTAL row ──
        total_errors        = sum(col_error_counts.values())
        total_record_count  = total_rows * len(FIELD_ORDER)
        total_pct_error     = round((total_errors / total_record_count) * 100, 2) if total_record_count else 0
        total_pct_health    = round(100 - total_pct_error, 2)

        ws.cell(row=row_num, column=1, value="")
        ws.cell(row=row_num, column=2, value="TOTAL")
        ws.cell(row=row_num, column=3, value=total_errors)
        ws.cell(row=row_num, column=4, value=total_record_count)
        ws.cell(row=row_num, column=5, value=f"{total_pct_health}%")
        ws.cell(row=row_num, column=6, value=f"{total_pct_error}%")
        ws.cell(row=row_num, column=7, value="")

        for c in range(1, 8):
            ws.cell(row=row_num, column=c).font      = Font(name="Arial", bold=True, size=10)
            ws.cell(row=row_num, column=c).fill      = TOTAL_FILL
            ws.cell(row=row_num, column=c).border    = THIN_BORDER
            ws.cell(row=row_num, column=c).alignment = Alignment(horizontal="center", vertical="center")

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
        col_widths = [6, 48, 14, 16, 12, 12, 90]
        for c_idx, width in enumerate(col_widths, start=1):
            ws.column_dimensions[get_column_letter(c_idx)].width = width

    # ── Per-field error sheets  (ALL source columns) ──
    def _write_field_error_sheets(self, wb, df: pd.DataFrame, all_cols: list):
        field_errors = self.validator.get_errors_by_field()

        for field_name in FIELD_ORDER:
            if field_name not in field_errors:
                continue

            row_indices = field_errors[field_name]
            sheet_name  = field_name[:31].replace("/", "-").replace("\\", "-").replace("*", "")
            ws          = wb.create_sheet(sheet_name)

            keep_here = all_cols + ["ERROR_COLUMNS"]
            subset    = df.loc[row_indices, keep_here].copy()

            field_err_series        = self.validator.get_field_error_series(field_name)
            subset["ERROR_COLUMNS"] = subset.index.map(
                lambda i: field_err_series.get(i, "")
            )

            self._write_header(ws, subset.columns)
            col_idx_map = {col: i for i, col in enumerate(subset.columns, start=1)}

            # MATERIAL sheet highlights MATERIAL + PLANT (the Part-Site combo)
            if field_name == "MATERIAL":
                highlight_cols = (PS_MATERIAL_COL, PS_PLANT_COL)
            else:
                highlight_cols = (field_name,)

            for excel_row, (orig_idx, row_data) in enumerate(subset.iterrows(), start=2):
                for c_idx, (col, value) in enumerate(zip(subset.columns, row_data), start=1):
                    cell           = ws.cell(row=excel_row, column=c_idx, value=value)
                    cell.font      = BODY_FONT
                    cell.alignment = Alignment(vertical="center")
                    cell.fill      = WHITE_FILL
                    cell.border    = THIN_BORDER

                for involved_col in highlight_cols:
                    if involved_col in col_idx_map:
                        target_cell      = ws.cell(row=excel_row, column=col_idx_map[involved_col])
                        target_cell.fill = RED_FILL
                        target_cell.font = ERR_FONT

            self._set_widths(ws)
            ws.freeze_panes = "A2"

            note_row = len(subset) + 3
            ws.cell(
                row=note_row, column=1,
                value=f"Total error rows for '{field_name}': {len(subset)}",
            ).font = Font(name="Arial", italic=True, size=9, bold=True)

    # ── Rules sheet ───────────────────────────
    def _write_rules_sheet(self, wb):
        ws = wb.create_sheet(self.SHEET_RULES)

        ws.merge_cells("A1:C1")
        title_cell           = ws.cell(row=1, column=1,
                                        value="PartSource(Make) Table – Business Validation Rules")
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

        for field in FIELD_ORDER:
            rules_list = self.RULES_CONTENT.get(field, [])
            num_rules  = len(rules_list)

            for r_idx, rule_text in enumerate(rules_list):
                num_cell           = ws.cell(row=current_row, column=1,
                                             value=rule_num if r_idx == 0 else "")
                num_cell.font      = Font(name="Arial", size=10, bold=(r_idx == 0))
                num_cell.fill      = RULE_FILL
                num_cell.border    = THIN_BORDER
                num_cell.alignment = Alignment(horizontal="center", vertical="center")

                field_cell           = ws.cell(row=current_row, column=2,
                                               value=field if r_idx == 0 else "")
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
        ws.column_dimensions["B"].width = 45
        ws.column_dimensions["C"].width = 100

    # ── Main write ────────────────────────────
    def write(self):
        v  = self.validator
        df = v.df.copy()

        # All source columns, in original order — used for per-field error sheets
        all_cols = list(df.columns)

        error_series        = v.get_error_series()
        df["ERROR_COLUMNS"] = df.index.map(
            lambda i: error_series.get(i, "") if i in error_series.index else ""
        )

        wb               = Workbook()
        ws_summary       = wb.active
        ws_summary.title = self.SHEET_SUMMARY
        self._write_summary_sheet_into(ws_summary, v.error_map, total_rows=len(df))

        self._write_rules_sheet(wb)
        self._write_field_error_sheets(wb, df, all_cols)

        wb.save(self.output_path)

        fields_with_errors = [f for f in FIELD_ORDER if f in v.get_errors_by_field()]
        print(f"\n✅  Output saved  → {self.output_path}")
        print(f"   Total rows                          : {len(df)}")
        print(f"   Error rows                          : {len(v.error_map)}")
        for sub_label in MATERIAL_SUB_CATEGORIES:
            print(f"     MATERIAL - {sub_label:<38}: {v.material_sub_counts[sub_label]}")
        print(f"   Field sheets                        : {fields_with_errors}")


# ══════════════════════════════════════════════
#  Orchestrator
# ══════════════════════════════════════════════
class PartSourceMakeBusinessTableProcessor:

    def __init__(self, ps_make_path: str, bom_path: str, source_constraint_path: str,
                 partmaster_path: str, output_path: str):
        self.validator = PartSourceMakeBusinessTableValidator(
            ps_make_path, bom_path, source_constraint_path, partmaster_path
        )
        self.writer    = PartSourceMakeBusinessReportWriter(self.validator, output_path)

    def run(self):
        print("📂  Loading files …")
        self.validator.load()
        print(f"    PartSource(Make) columns detected : {list(self.validator.df.columns)}")
        print(f"    Valid BOM Part-Site combinations   : {len(self.validator.valid_bom_combos)}")
        print(f"    Valid SourceConstraint combinations: {len(self.validator.valid_sc_combos)}")
        if self.validator.sc_factor_lookup is not None:
            print(f"    Source Constraint factor lookup    : {len(self.validator.sc_factor_lookup)}")
        if self.validator.partmaster_lookup is not None:
            print(f"    Part master Part-Site combinations : {len(self.validator.partmaster_lookup)}")
        print("🔍  Validating business rules …")
        self.validator.validate()
        print("📝  Writing report …")
        self.writer.write()


# ══════════════════════════════════════════════
#  Entry Point
# ══════════════════════════════════════════════
if __name__ == "__main__":
    processor = PartSourceMakeBusinessTableProcessor(
        ps_make_path           = PARTSOURCE_MAKE_INPUT_FILE,
        bom_path               = BOM_INPUT_FILE,
        source_constraint_path = SOURCE_CONSTRAINT_INPUT_FILE,
        partmaster_path        = PARTMASTER_INPUT_FILE,
        output_path            = OUTPUT_FILE,
    )
    processor.run()
