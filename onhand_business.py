import re
import traceback
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ─────────────────────────────────────────────
#  FILE PATHS
# ─────────────────────────────────────────────
ONHAND_INPUT_FILE = r"C:\Users\SW526XH\Downloads\Go Live-2\GL2 Extracts\OnHand.tab"
PART_INPUT_FILE   = r"C:\Users\SW526XH\Downloads\Go Live-2\GL2 Extracts\Part.tab"   # <-- check this path
OUTPUT_FILE       = r"C:\Users\SW526XH\Downloads\Go Live-2\Onhand\Validated_Onhand_Business.xlsx"


# ─────────────────────────────────────────────
#  DATE FORMAT
# ─────────────────────────────────────────────
DATE_PATTERN = re.compile(r'^\d{4}(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])$')  # YYYYMMDD


# ─────────────────────────────────────────────
#  PART CLASS (FG vs RM/PM/SFG) CONFIGURATION
#  Onhand has MATERIALNUMBER / MATERIALTYPE but no IBPSTATUS, so IBPSTATUS and
#  MRPTYPE are looked up in the Part extract on MATERIALNUMBER + PLANT.
# ─────────────────────────────────────────────
PART_FG      = "FG"
PART_RMPMSFG = "RM/PM/SFG"

# Part extract columns used for the lookup
PART_KEY_COLS    = ["MATERIALNUMBER", "PLANT"]
PART_LOOKUP_COLS = ["IBPSTATUS", "MRPTYPE"]

# FG
FG_FERT_HAWA_TYPES  = {"FERT", "HAWA"}          # FG only when IBPSTATUS in FG_IBP_STATUSES
FG_IBP_STATUSES     = {"IBP", "IBPEXP"}         # normalised (no spaces / underscores)
FG_ANY_IBP_TYPES    = {"UNBW"}                  # FG for any IBPSTATUS (even if not found in Part)

# RM / PM / SFG
HAWA_RM_IBP_STATUS  = "N"                       # HAWA with IBPSTATUS N  -> RM/PM/SFG
MRPTYPE_TYPES       = {"HIBE", "RAWT"}          # RM/PM/SFG only when Part MRPTYPE == "X0"
MRPTYPE_VALUE       = "X0"
RMPMSFG_AS_IS_TYPES = {"BLND", "HALB", "ROH", "VERP", "BASB"}   # taken as they are


# ─────────────────────────────────────────────
#  NETTABLE / NON-NETTABLE CONFIGURATION
# ─────────────────────────────────────────────
NETTABLE_FIELD = "TYPE_SUBTYPE_STORAGELOCATION"     # logical field name for this rule
NETTABLE_COLS  = ["MATERIALTYPE", "TYPE", "SUBTYPE", "STORAGELOCATION"]

NETTABLE       = "NETTABLE"
NON_NETTABLE   = "NONNETTABLE"

# TYPE -> allowed SUBTYPE values (normalised: upper-case, no spaces/underscores/hyphens)
# Part FG
TYPE_RULES_FG = {
    "BLOCKEDSTOCK":                    {NON_NETTABLE},
    "BLOCKEDSTOCKRETURNS":             {NON_NETTABLE},
    "QUALITYINSPECTION":               {NETTABLE, NON_NETTABLE},
    "SCHEDULEDFORDELIVERY":            {NON_NETTABLE},
    "VALUATEDUNRESTRICTEDUSESTOCK":    {NETTABLE, NON_NETTABLE},
    "TENTATIVEINVENTORYPOSTINGSHIFTS": {NON_NETTABLE},
    "VENDORSTOCK":                     {NETTABLE},
}

# Part RM / PM / SFG
TYPE_RULES_RMPMSFG = {
    "BLOCKEDSTOCK":                    {NON_NETTABLE},
    "BLOCKEDSTOCKRETURNS":             {NON_NETTABLE},
    "QUALITYINSPECTION":               {NETTABLE},
    "SCHEDULEDFORDELIVERY":            {NON_NETTABLE},
    "VALUATEDUNRESTRICTEDUSESTOCK":    {NETTABLE},
    "VENDORSTOCK":                     {NETTABLE},
}

PART_CLASS_RULES = {
    PART_FG:      TYPE_RULES_FG,
    PART_RMPMSFG: TYPE_RULES_RMPMSFG,
}

# Nettable STORAGELOCATION sub-rule applies only to these part classes
LOCATION_CHECK_PART_CLASSES = {PART_FG}

# Blank location is represented by this token
BLANK_LOCATION_TOKEN = "BLANK LOCATION"

NETTABLE_LOCATIONS = {
    BLANK_LOCATION_TOKEN,
    "0007", "0003", "0009", "0036", "0037", "0040",
    "1809", "1817", "1879", "FG01",
}


# ─────────────────────────────────────────────
#  VALUATEDUNRESTRICTEDUSESTOCK (quantity) RULE
# ─────────────────────────────────────────────
QTY_FIELD        = "VALUATEDUNRESTRICTEDUSESTOCK"   # rule / field name shown in the report
QTY_COL_FALLBACKS = ["VALUATEDUNRESTRICTEDUSESTOCK", "QTY"]   # source column names to look for


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
SUB_FILL        = PatternFill("solid", start_color="FFFFFF", end_color="FFFFFF")

HDR_FONT    = Font(bold=True, name="Arial")
BODY_FONT   = Font(name="Arial", size=10)
ERR_FONT    = Font(name="Arial", size=10, bold=True, color="FFFFFF")
THIN_BORDER = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"),  bottom=Side(style="thin"),
)

# ─────────────────────────────────────────────
#  FIELD ORDER  (only fields with business rules)
# ─────────────────────────────────────────────
FIELD_ORDER = ["DATEOFLASTGOODSRECEIPT", NETTABLE_FIELD, QTY_FIELD]

# Field(s) that use sub-rows in the summary
FIELDS_WITH_SUB_ROWS = {"DATEOFLASTGOODSRECEIPT", NETTABLE_FIELD}

# Source columns highlighted red on the error sheet for each rule
# (QTY_FIELD is resolved at run time to the actual source column name)
INVOLVED_COLUMNS = {
    "DATEOFLASTGOODSRECEIPT": ["DATEOFLASTGOODSRECEIPT"],
    NETTABLE_FIELD:           NETTABLE_COLS,
}

# Summary sub-rows:  field -> [(label, key, summary reason, substring marker in row reason)]
SUB_ROW_DEFS = {
    "DATEOFLASTGOODSRECEIPT": [
        ("  ↳ Blank Field", "blank",
         "DATEOFLASTGOODSRECEIPT: Field is blank",
         "Field is blank"),
        ("  ↳ Invalid Format (not YYYYMMDD)", "bad_format",
         "DATEOFLASTGOODSRECEIPT: Does not follow required format YYYYMMDD",
         "does not follow required format"),
        # ("  ↳ Duplicate Value", "duplicate",
        #  "DATEOFLASTGOODSRECEIPT: Value is not unique (duplicate found)",
        #  "duplicate value"),
    ],
    NETTABLE_FIELD: [
        ("  ↳ Unrecognised TYPE for Part Class", "unknown_type",
         "TYPE: Value is not valid for the given Part Class (FG / RM-PM-SFG)",
         "is not a recognised Onhand TYPE"),
        ("  ↳ Invalid SUBTYPE for TYPE", "bad_subtype",
         "SUBTYPE: Value not allowed for the given Part Class / TYPE (Nettable / Non Nettable)",
         "is not an allowed SUBTYPE for TYPE"),
        ("  ↳ Nettable – Invalid STORAGELOCATION (FG)", "n_location",
         "STORAGELOCATION: Not in the allowed Nettable location list",
         "is not a valid Nettable location"),
    ],
}


# ══════════════════════════════════════════════
#  Helpers
# ══════════════════════════════════════════════
def _is_blank(value) -> bool:
    return pd.isna(value) or str(value).strip() == ""


def _norm_key(value) -> str:
    """Upper-case and strip spaces / underscores / hyphens (for TYPE / SUBTYPE / IBPSTATUS matching)."""
    if pd.isna(value):
        return ""
    return re.sub(r"[\s_\-]+", "", str(value)).upper()


def _norm_id(value) -> str:
    """Upper-case + trim (for MATERIALNUMBER / PLANT / MRPTYPE matching)."""
    if pd.isna(value):
        return ""
    return str(value).strip().upper()


# ══════════════════════════════════════════════
#  Rule Engine
# ══════════════════════════════════════════════
class OnhandBusinessRuleEngine:
    """
    Rules:
      DATEOFLASTGOODSRECEIPT
        a) Field should not be blank.
        b) Field should be in the format YYYYMMDD.
        (Uniqueness check is currently disabled.)

      TYPE_SUBTYPE_STORAGELOCATION
        Part class is derived first (Part extract lookup on MATERIALNUMBER + PLANT):
          FG        : MATERIALTYPE FERT/HAWA with IBPSTATUS IBP/IBPEXP, or UNBW (any IBPSTATUS)
          RM/PM/SFG : HAWA with IBPSTATUS N; HIBE/RAWT with MRPTYPE X0;
                      BLND/HALB/ROH/VERP/BASB as they are
        Rows that match neither class are not checked.
        a) TYPE must be recognised for the Part class.
        b) SUBTYPE must be allowed for that Part class + TYPE.
        c) FG only: if SUBTYPE is Nettable, STORAGELOCATION must be in the
           Nettable location list (blank location is valid).

      VALUATEDUNRESTRICTEDUSESTOCK
        Field should be > 0 (blank values are skipped).
    """

    def __init__(self, part_lookup: dict, qty_col, check_nettable: bool = True):
        self.part_lookup    = part_lookup     # (MATERIALNUMBER, PLANT) -> (IBPSTATUS, MRPTYPE)
        self.qty_col        = qty_col         # actual source column for the quantity rule (or None)
        self.check_nettable = check_nettable
        self.class_counts   = {PART_FG: 0, PART_RMPMSFG: 0, "UNCLASSIFIED": 0}

    # ── Part class derivation ─────────────────
    def _part_class(self, mat_key: str, mat_num: str, plant: str):
        if mat_key in FG_ANY_IBP_TYPES:
            return PART_FG

        if mat_key in RMPMSFG_AS_IS_TYPES:
            return PART_RMPMSFG

        if mat_key in FG_FERT_HAWA_TYPES or mat_key in MRPTYPE_TYPES:
            ibp, mrp = self.part_lookup.get((mat_num, plant), ("", ""))

            if mat_key in FG_FERT_HAWA_TYPES:
                if ibp in FG_IBP_STATUSES:
                    return PART_FG
                if mat_key == "HAWA" and ibp == HAWA_RM_IBP_STATUS:
                    return PART_RMPMSFG
                return None

            # HIBE / RAWT
            if mrp == MRPTYPE_VALUE:
                return PART_RMPMSFG
            return None

        return None

    # ── DATEOFLASTGOODSRECEIPT ────────────────
    def _check_dlgr(self, row, reasons: dict):
        val = row.get("DATEOFLASTGOODSRECEIPT", "")

        # ── Blank check ──
        if _is_blank(val):
            reasons["DATEOFLASTGOODSRECEIPT"] = "DATEOFLASTGOODSRECEIPT: Field is blank"
            return

        val_str     = str(val).strip()
        sub_reasons = []

        # ── Format check ──
        if not DATE_PATTERN.match(val_str):
            sub_reasons.append(
                f"DATEOFLASTGOODSRECEIPT: '{val_str}' does not follow required format YYYYMMDD"
            )

        # ── Uniqueness check (disabled) ──
        # occurrences = self.value_counts.get(val_str, 0)
        # if occurrences > 1:
        #     sub_reasons.append(
        #         f"DATEOFLASTGOODSRECEIPT: '{val_str}' is a duplicate value (appears {occurrences} times)"
        #     )

        if sub_reasons:
            reasons["DATEOFLASTGOODSRECEIPT"] = " | ".join(sub_reasons)

    # ── VALUATEDUNRESTRICTEDUSESTOCK > 0 ──────
    def _check_qty(self, row, reasons: dict):
        if not self.qty_col:
            return

        val = row.get(self.qty_col, "")

        # Blank values are allowed for this rule
        if _is_blank(val):
            return

        try:
            num = float(str(val).strip().replace(",", ""))
        except (ValueError, TypeError):
            return   # non-numeric values are not flagged by this rule

        if num <= 0:
            reasons[QTY_FIELD] = f"{QTY_FIELD}: '{str(val).strip()}' should be > 0"

    # ── TYPE / SUBTYPE / STORAGELOCATION ──────
    def _check_nettable(self, row, reasons: dict):
        if not self.check_nettable:
            return

        mat_key      = _norm_id(row.get("MATERIALTYPE", ""))
        mat_num      = _norm_id(row.get("MATERIALNUMBER", ""))
        plant        = _norm_id(row.get("PLANT", ""))
        type_raw     = row.get("TYPE", "")
        subtype_raw  = row.get("SUBTYPE", "")
        loc_raw      = row.get("STORAGELOCATION", "")

        type_disp    = "" if pd.isna(type_raw)    else str(type_raw).strip()
        subtype_disp = "" if pd.isna(subtype_raw) else str(subtype_raw).strip()
        loc_disp     = "" if pd.isna(loc_raw)     else str(loc_raw).strip()

        # ── Step 0: derive Part class; rows that fit neither class are not checked ──
        part_class = self._part_class(mat_key, mat_num, plant)
        if part_class is None:
            self.class_counts["UNCLASSIFIED"] += 1
            return
        self.class_counts[part_class] += 1

        type_rules = PART_CLASS_RULES[part_class]

        # ── Sub-rule a: TYPE must be recognised for this Part class ──
        type_key = _norm_key(type_disp)
        if type_key not in type_rules:
            shown = type_disp if type_disp else "<blank>"
            reasons[NETTABLE_FIELD] = (
                f"TYPE: '{shown}' is not a recognised Onhand TYPE for Part class '{part_class}' "
                f"(MATERIALTYPE '{mat_key}')"
            )
            return

        # ── Sub-rule b: SUBTYPE must be allowed for this Part class + TYPE ──
        subtype_key = _norm_key(subtype_disp)
        allowed     = type_rules[type_key]
        if subtype_key not in allowed:
            shown_sub   = subtype_disp if subtype_disp else "<blank>"
            allowed_txt = " / ".join(
                "Nettable" if a == NETTABLE else "Non Nettable" for a in sorted(allowed)
            )
            reasons[NETTABLE_FIELD] = (
                f"SUBTYPE: '{shown_sub}' is not an allowed SUBTYPE for TYPE '{type_disp}' "
                f"with Part class '{part_class}' (MATERIALTYPE '{mat_key}') (allowed: {allowed_txt})"
            )
            return

        # ── Sub-rule c: Nettable stock must sit in a valid Nettable location (FG only) ──
        if part_class in LOCATION_CHECK_PART_CLASSES and subtype_key == NETTABLE:
            loc_key = loc_disp.upper() if loc_disp else BLANK_LOCATION_TOKEN
            if loc_key not in NETTABLE_LOCATIONS:
                shown_loc = loc_disp if loc_disp else "<blank>"
                reasons[NETTABLE_FIELD] = (
                    f"STORAGELOCATION: '{shown_loc}' is not a valid Nettable location "
                    f"(Part class '{part_class}', MATERIALTYPE '{mat_key}', TYPE '{type_disp}')"
                )

    # ── Row entry point ──
    def validate_row(self, row) -> dict:
        reasons = {}
        self._check_dlgr(row, reasons)
        self._check_nettable(row, reasons)
        self._check_qty(row, reasons)
        return reasons


# ══════════════════════════════════════════════
#  Validator
# ══════════════════════════════════════════════
class OnhandBusinessTableValidator:

    def __init__(self, onhand_path: str, part_path: str):
        self.onhand_path  = onhand_path
        self.part_path    = part_path
        self.df           = pd.DataFrame()
        self.part_lookup  = {}
        self.qty_col      = None
        self.class_counts = {}
        self.error_map    = {}   # row_idx -> [failed field names]
        self.reason_map   = {}   # row_idx -> {field: reason}

    def load(self):
        self.df = pd.read_csv(self.onhand_path, sep="\t", dtype=str)
        self.df.columns = [c.strip().upper() for c in self.df.columns]

    # ── Part extract lookup: (MATERIALNUMBER, PLANT) -> (IBPSTATUS, MRPTYPE) ──
    def _load_part_lookup(self):
        try:
            part = pd.read_csv(self.part_path, sep="\t", dtype=str)
        except Exception as exc:
            print(f"    ⚠️  Could not read Part extract ({self.part_path}): {exc}")
            return {}, False

        part.columns = [c.strip().upper() for c in part.columns]
        needed  = PART_KEY_COLS + PART_LOOKUP_COLS
        missing = [c for c in needed if c not in part.columns]
        if missing:
            print(f"    ⚠️  Part extract is missing column(s): {missing}")
            return {}, False

        lookup = {}
        for mat, plant, ibp, mrp in zip(
            part["MATERIALNUMBER"], part["PLANT"], part["IBPSTATUS"], part["MRPTYPE"]
        ):
            key = (_norm_id(mat), _norm_id(plant))
            if key not in lookup:                      # first occurrence wins
                lookup[key] = (_norm_key(ibp), _norm_id(mrp))
        print(f"    Part lookup built       : {len(lookup)} MATERIALNUMBER+PLANT keys")
        return lookup, True

    def validate(self):
        # Quantity column (VALUATEDUNRESTRICTEDUSESTOCK, falls back to QTY)
        self.qty_col = next((c for c in QTY_COL_FALLBACKS if c in self.df.columns), None)
        if self.qty_col is None:
            print(f"    ⚠️  {QTY_FIELD} rule skipped – none of {QTY_COL_FALLBACKS} found in Onhand")
        else:
            print(f"    {QTY_FIELD} rule reads column : {self.qty_col}")

        # Columns needed for the Nettable rule
        required_onhand = ["MATERIALNUMBER", "PLANT"] + NETTABLE_COLS
        missing = [c for c in required_onhand if c not in self.df.columns]
        if missing:
            print(f"    ⚠️  Nettable/Location rule skipped – missing Onhand column(s): {missing}")

        part_ok = False
        if not missing:
            self.part_lookup, part_ok = self._load_part_lookup()
            if not part_ok:
                print("    ⚠️  Nettable/Location rule skipped – Part lookup unavailable")

        engine = OnhandBusinessRuleEngine(
            self.part_lookup, self.qty_col, check_nettable=(not missing and part_ok)
        )

        first_error_shown = False
        for idx, row in self.df.iterrows():
            try:
                reasons = engine.validate_row(row)
            except Exception:
                reasons = {}
                if not first_error_shown:
                    first_error_shown = True
                    print(f"    ⚠️  Row {idx} raised an exception (further ones suppressed):")
                    traceback.print_exc()

            if reasons:
                self.error_map[idx]  = list(reasons.keys())
                self.reason_map[idx] = reasons

        self.class_counts = engine.class_counts

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

    # ── Sub-count helper (generic, substring based) ──
    def get_field_subcounts(self, field_name: str) -> dict:
        defs   = SUB_ROW_DEFS.get(field_name, [])
        counts = {key: 0 for _, key, _, _ in defs}
        for col_reason in self.reason_map.values():
            reason = col_reason.get(field_name, "")
            if not reason:
                continue
            for _, key, _, marker in defs:
                if marker in reason:
                    counts[key] += 1
        return counts


# ══════════════════════════════════════════════
#  Report Writer
# ══════════════════════════════════════════════
class OnhandBusinessReportWriter:

    SHEET_SUMMARY = "Summary"
    SHEET_RULES   = "Rules"

    RULES_CONTENT = {
        "DATEOFLASTGOODSRECEIPT": [
            "Field must not be blank.",
            # "Values must be unique — no duplicates allowed.",   # disabled
            "Field must follow the format: YYYYMMDD.",
        ],
        NETTABLE_FIELD: [
            "Part class is derived using the Part extract (lookup on MATERIALNUMBER + PLANT). "
            "FG: MATERIALTYPE FERT / HAWA with IBPSTATUS IBP or IBPEXP, or UNBW (any IBPSTATUS). "
            "RM/PM/SFG: HAWA with IBPSTATUS N; HIBE / RAWT with MRPTYPE X0; "
            "BLND, HALB, ROH, VERP, BASB taken as they are. Other rows are not checked.",
            "Part FG: TYPE/SUBTYPE must be valid — BLOCKEDSTOCK, BLOCKEDSTOCKRETURNS, "
            "SCHEDULEDFORDELIVERY, TENTATIVEINVENTORYPOSTINGSHIFTS: Non Nettable only; "
            "QUALITYINSPECTION, VALUATEDUNRESTRICTEDUSESTOCK: Nettable or Non Nettable; "
            "VENDORSTOCK: Nettable only.",
            "Part FG: If SUBTYPE is Nettable, STORAGELOCATION must be one of: "
            + ", ".join(["Blank location"] + sorted(NETTABLE_LOCATIONS - {BLANK_LOCATION_TOKEN})) + ".",
            "Part RM/PM/SFG: TYPE/SUBTYPE must be valid — BLOCKEDSTOCK, BLOCKEDSTOCKRETURNS, "
            "SCHEDULEDFORDELIVERY: Non Nettable only; QUALITYINSPECTION, "
            "VALUATEDUNRESTRICTEDUSESTOCK, VENDORSTOCK: Nettable only. "
            "(TENTATIVEINVENTORYPOSTINGSHIFTS is not valid for RM/PM/SFG.)",
        ],
        QTY_FIELD: [
            "Field should be > 0.",
        ],
    }

    def __init__(self, validator: OnhandBusinessTableValidator, output_path: str):
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
                                italic: bool = False):
        for c in range(1, num_cols + 1):
            cell           = ws.cell(row=row_num, column=c)
            cell.font      = Font(name="Arial", bold=bold, italic=italic, size=10)
            cell.border    = THIN_BORDER
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if fill:
                cell.fill = fill

    # ══════════════════════════════════════════
    #  Summary sheet  (only fields that have business rules)
    # ══════════════════════════════════════════
    def _write_summary_sheet_into(self, ws, error_map: dict, total_rows: int):

        # ── Row 1 : Title ──
        ws.merge_cells("A1:G1")
        title_cell           = ws.cell(row=1, column=1, value="Onhand Business Validation Summary")
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

        # ── Build per-field error counts ──
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

            ws.cell(row=row_num, column=1, value=field_num)
            ws.cell(row=row_num, column=2, value=col_name)
            ws.cell(row=row_num, column=3, value=count)
            ws.cell(row=row_num, column=4, value=total_rows)
            ws.cell(row=row_num, column=5, value=f"{pct_health}%")
            ws.cell(row=row_num, column=6, value=f"{pct_error}%")
            ws.cell(row=row_num, column=7, value="")

            self._style_summary_data_row(ws, row_num, fill=WHITE_FILL)
            ws.cell(row=row_num, column=7).alignment = Alignment(
                horizontal="left", vertical="center", wrap_text=True
            )
            row_num += 1

            if col_name in FIELDS_WITH_SUB_ROWS and has_errors:
                subcounts = self.validator.get_field_subcounts(col_name)
                for sub_label, sub_key, sub_reason, _marker in SUB_ROW_DEFS[col_name]:
                    sub_count      = subcounts.get(sub_key, 0)
                    sub_pct_err    = round((sub_count / total_rows) * 100, 2) if total_rows else 0
                    sub_pct_health = round(100 - sub_pct_err, 2)
                    ws.cell(row=row_num, column=1, value="")
                    ws.cell(row=row_num, column=2, value=sub_label)
                    ws.cell(row=row_num, column=3, value=sub_count)
                    ws.cell(row=row_num, column=4, value=total_rows)
                    ws.cell(row=row_num, column=5, value=f"{sub_pct_health}%")
                    ws.cell(row=row_num, column=6, value=f"{sub_pct_err}%")
                    ws.cell(row=row_num, column=7, value=sub_reason)
                    self._style_summary_data_row(ws, row_num, fill=SUB_FILL, italic=True)
                    ws.cell(row=row_num, column=2).alignment = Alignment(
                        horizontal="left", vertical="center", indent=1
                    )
                    ws.cell(row=row_num, column=7).alignment = Alignment(
                        horizontal="left", vertical="center", wrap_text=True
                    )
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
        col_widths = [6, 46, 14, 16, 12, 12, 70]
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

            # Columns to highlight red for THIS sheet's rule
            if field_name == QTY_FIELD:
                highlight_cols = [self.validator.qty_col] if self.validator.qty_col else []
            else:
                highlight_cols = INVOLVED_COLUMNS.get(field_name, [field_name])

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
        title_cell           = ws.cell(row=1, column=1, value="Onhand Table – Business Validation Rules")
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
        ws.column_dimensions["C"].width = 95

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
        print(f"   Total rows    : {len(df)}")
        print(f"   Error rows    : {len(v.error_map)}")
        print(f"   Field sheets  : {fields_with_errors}")
        print(f"   Part class split (rows checked by the Nettable rule): {v.class_counts}")


# ══════════════════════════════════════════════
#  Orchestrator
# ══════════════════════════════════════════════
class OnhandBusinessTableProcessor:

    def __init__(self, onhand_path: str, part_path: str, output_path: str):
        self.validator = OnhandBusinessTableValidator(onhand_path, part_path)
        self.writer    = OnhandBusinessReportWriter(self.validator, output_path)

    def run(self):
        print("📂  Loading file …")
        self.validator.load()
        print(f"    Onhand columns detected : {list(self.validator.df.columns)}")
        print("🔍  Validating business rules …")
        self.validator.validate()
        print("📝  Writing report …")
        self.writer.write()


# ══════════════════════════════════════════════
#  Entry Point
# ══════════════════════════════════════════════
if __name__ == "__main__":
    processor = OnhandBusinessTableProcessor(
        onhand_path = ONHAND_INPUT_FILE,
        part_path   = PART_INPUT_FILE,
        output_path = OUTPUT_FILE,
    )
    processor.run()
