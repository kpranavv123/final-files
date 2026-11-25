import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ─────────────────────────────────────────────
#  FILE PATHS
# ─────────────────────────────────────────────
PART_INPUT_FILE              = r"C:\Users\SW526XH\Downloads\Go Live-2\GL2 Extracts\PARTSITE_2026-11-13-1920.tab"
PARTSOURCE_MAKE_INPUT_FILE   = r"C:\Users\SW526XH\Downloads\Go Live-2\GL2 Extracts\PartSource(Make)2026-11-13-1930.tsv"
PARTSOURCE_BUY_INPUT_FILE    = r"C:\Users\SW526XH\Downloads\Go Live-2\GL2 Extracts\PartSource_Buy_2026-11-13-1935.tsv"
SOURCE_CONSTRAINT_INPUT_FILE = r"C:\Users\SW526XH\Downloads\Go Live-2\GL2 Extracts\srcconstraint_2026-11-13-1929.tsv"
OUTPUT_FILE                  = r"C:\Users\SW526XH\Downloads\Go Live-2\Part_RMPM\Validated_Part_Business.xlsx"


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

HDR_FONT    = Font(bold=True, name="Arial")
BODY_FONT   = Font(name="Arial", size=10)
ERR_FONT    = Font(name="Arial", size=10, bold=True, color="FFFFFF")
THIN_BORDER = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"),  bottom=Side(style="thin"),
)


# ─────────────────────────────────────────────
#  COLUMN CONFIG
# ─────────────────────────────────────────────
RULE_FIELD = "PROCUREMENTTYPE"

# Part extract
PART_MATERIAL_COL   = "MATERIALNUMBER"
PART_PLANT_COL      = "PLANT"
PROCUREMENT_COL     = "PROCUREMENTTYPE"
PROCUREMENT_VALUE_E = "E"
PROCUREMENT_VALUE_F = "F"

# PartSource (make) extract
PS_MATERIAL_COL = "MATERIAL"
PS_PLANT_COL    = "PLANT"

# PartSource (buy) extract
PB_MATERIAL_COL = "MATERIALNUMBER"
PB_PLANT_COL    = "PLANT"

# Source Constraint extract
SC_MATERIAL_COL = "MATERIAL"
SC_PLANT_COL    = "WORKCENTERPLANT"
SC_FACTOR_COL   = "CONSTRAINTFACTOR"

# ── Field rules ──
MINSHELFLIFE_COL    = "MINREMSHELFLIFE"          # must be > 0 when BATCHMANAGED = "Y"
BATCHMANAGED_COL    = "BATCHMANAGED"             # NEW: Part extract column driving the MINREMSHELFLIFE rule
BATCHMANAGED_YES    = "Y"                        # NEW
INTERVALSEXPIRY_COL = "INTERVALSTOEXPIRYDAYS"    # must be >= 0
STANDARDPRICE_COL   = "STANDARDPRICE"            # (a) non-zero for non FERT/HAWA, (b) non-zero for RM/PM/SFG part class

# Material-type column in the Part extract (first candidate found is used)
PRODUCT_TYPE_COL_CANDIDATES = ["PRODUCTYPE", "PRODUCTTYPE"]
FG_HAWA_TYPES               = {"FERT", "HAWA"}

# ── NEW: Part class derivation inputs (Part extract) ──
IBPSTATUS_COL = "IBPSTATUS"
MRPTYPE_COL   = "MRPTYPE"

# Part class derivation table (PRODUCTTYPE + IBPSTATUS + MRPTYPE -> Part Class)
#   "*" means any value (including blank).
#   X0 / XO are treated as the same value.
#
#   PRODUCTTYPE | IBPSTATUS        | MRPTYPE | Part Class
#   BLND        | *                | *       | SFG
#   HALB        | *                | *       | SFG
#   HAWA        | N                | *       | RM
#   HIBE        | blank or X0      | *       | PM
#   ROH         | *                | *       | RM
#   VERP        | *                | *       | PM
#   BASB        | *                | *       | SFG
#   RAWT        | *                | X0      | SFG
# Any other combination has no part class and is not checked by sub-rule (b).
PART_CLASS_SCOPE = ("RM", "PM", "SFG")

# Blank handling for the numeric rules (INTERVALSTOEXPIRYDAYS, STANDARDPRICE sub-rule a):
#   False -> blank values are skipped (not flagged)       [suite convention]
#   True  -> blank values are flagged as errors too
# Non-numeric (non-blank) values are always flagged.
# NOTE: STANDARDPRICE sub-rule (b) and MINREMSHELFLIFE (when BATCHMANAGED = Y)
#       always flag blanks, because the value must be > 0 / non-zero.
BLANK_FAILS_NUMERIC_CHECKS = False

# Sub-rules (b) and (c) behaviour when a Part-Site has NO row at all in the
# Source Constraint extract:
#   False -> not flagged (rule only fails when a matching row exists with factor 0)
#   True  -> flagged as an error as well
MISSING_IN_CONSTRAINT_IS_ERROR = False

# ── PROCUREMENTTYPE sub-rule ids ──
SUB_A = "a"
SUB_B = "b"
SUB_C = "c"
PROCUREMENT_SUBS = (SUB_A, SUB_B, SUB_C)

SUB_RULE_REASON = {
    SUB_A: "a) PROCUREMENTTYPE is 'E' but MATERIALNUMBER-PLANT combination is not present "
           "in PartSource (Make) extract (MATERIAL-PLANT)",
    SUB_B: "b) PROCUREMENTTYPE is 'E' but CONSTRAINTFACTOR is '0' for the MATERIAL-PLANT "
           "combination in Source Constraint extract",
    SUB_C: "c) PROCUREMENTTYPE is 'F' but MATERIALNUMBER-PLANT combination is not present "
           "in PartSource (Buy) extract (MATERIALNUMBER-PLANT)",
}

# ── NEW: STANDARDPRICE sub-rule ids ──
STD_SUBS = (SUB_A, SUB_B)
STD_SUB_REASON = {
    SUB_A: "a) STANDARDPRICE must be non-zero for rows where the product type is not FERT or HAWA",
    SUB_B: "b) STANDARDPRICE must be non-zero for rows whose part class is RM / PM / SFG "
           "(derived from PRODUCTTYPE + IBPSTATUS + MRPTYPE)",
}

# Fields that use sub-rows in the Summary: field -> (sub ids, reason dict)
FIELDS_WITH_SUBS = {
    RULE_FIELD:        (PROCUREMENT_SUBS, SUB_RULE_REASON),
    STANDARDPRICE_COL: (STD_SUBS,         STD_SUB_REASON),
}

# Reasons for the single-rule fields (no sub-rows in the Summary)
FIELD_REASON = {
    MINSHELFLIFE_COL:    "MINREMSHELFLIFE: must be a number greater than 0 when BATCHMANAGED is 'Y'.",
    INTERVALSEXPIRY_COL: "INTERVALSTOEXPIRYDAYS: must be a number greater than or equal to 0.",
}

FIELD_ORDER = [RULE_FIELD, MINSHELFLIFE_COL, INTERVALSEXPIRY_COL, STANDARDPRICE_COL]


# ══════════════════════════════════════════════
#  Rule Engine
# ══════════════════════════════════════════════
class PartBusinessRuleEngine:
    """
    Rule: PROCUREMENTTYPE
      a) 'E' Part-Site must be present in the PartSource (Make) extract (MATERIAL, PLANT).
      b) 'E' Part-Site must not have CONSTRAINTFACTOR '0' in the Source Constraint extract.
      c) 'F' Part-Site must be present in the PartSource (Buy) extract (MATERIALNUMBER, PLANT).

    Rule: MINREMSHELFLIFE     -> when BATCHMANAGED = 'Y' the value must be > 0           [CHANGED]
    Rule: INTERVALSTOEXPIRY   -> must be >= 0
    Rule: STANDARDPRICE
      a) non-zero when product type is not FERT/HAWA
      b) non-zero when part class is RM / PM / SFG                                       [NEW]
    """

    @staticmethod
    def _clean(series: pd.Series) -> pd.Series:
        return series.fillna("").astype(str).str.strip()

    @staticmethod
    def _parse_number(value):
        s = str(value).strip()
        if s == "" or s.lower() == "nan":
            return None
        try:
            return float(s)
        except ValueError:
            return None

    @staticmethod
    def _missing_cols(df: pd.DataFrame, cols: list) -> list:
        return [c for c in cols if c not in df.columns]

    def _proc_rows(self, part_df: pd.DataFrame, proc_value: str) -> pd.DataFrame:
        """Part rows with PROCUREMENTTYPE = proc_value and non-blank MATERIALNUMBER + PLANT.
        Returns a frame indexed like part_df with cleaned mat / plant columns."""
        proc  = self._clean(part_df[PROCUREMENT_COL]).str.upper()
        mat   = self._clean(part_df[PART_MATERIAL_COL])
        plant = self._clean(part_df[PART_PLANT_COL])
        mask  = (proc == proc_value.upper()) & (mat != "") & (plant != "")
        return pd.DataFrame({"mat": mat[mask], "plant": plant[mask]})

    # ── Generic "Part-Site must exist in another extract" check ──
    def _check_presence(self, part_df, other_df, other_mat_col, other_plant_col,
                        proc_value, sub_label, extract_name) -> dict:
        miss_part  = self._missing_cols(part_df, [PROCUREMENT_COL, PART_MATERIAL_COL, PART_PLANT_COL])
        miss_other = self._missing_cols(other_df, [other_mat_col, other_plant_col])
        if miss_part or miss_other:
            print(f"⚠️   Sub-rule ({sub_label}) skipped — missing column(s): "
                  f"Part={miss_part}, {extract_name}={miss_other}")
            return {}

        other_keys = set(zip(self._clean(other_df[other_mat_col]),
                             self._clean(other_df[other_plant_col])))

        rows   = self._proc_rows(part_df, proc_value)
        result = {}
        for idx, mat, plant in zip(rows.index, rows["mat"], rows["plant"]):
            if (mat, plant) not in other_keys:
                result[idx] = (
                    f"PROCUREMENTTYPE: '{proc_value}' but {PART_MATERIAL_COL} '{mat}' + "
                    f"{PART_PLANT_COL} '{plant}' is not present in {extract_name} extract"
                )
        return result

    # ── Sub-rule (a) ──
    def check_partsource(self, part_df: pd.DataFrame, ps_df: pd.DataFrame) -> dict:
        """E rows whose MATERIALNUMBER-PLANT is missing from the PartSource (Make) extract."""
        return self._check_presence(
            part_df, ps_df, PS_MATERIAL_COL, PS_PLANT_COL,
            PROCUREMENT_VALUE_E, SUB_A, "PartSource (Make)",
        )

    # ── Sub-rule (c) ──
    def check_partsource_buy(self, part_df: pd.DataFrame, pb_df: pd.DataFrame) -> dict:
        """F rows whose MATERIALNUMBER-PLANT is missing from the PartSource (Buy) extract."""
        return self._check_presence(
            part_df, pb_df, PB_MATERIAL_COL, PB_PLANT_COL,
            PROCUREMENT_VALUE_F, SUB_C, "PartSource (Buy)",
        )

    # ── Sub-rule (b) ──
    def check_constraint_factor(self, part_df: pd.DataFrame, sc_df: pd.DataFrame) -> dict:
        """E rows whose MATERIAL-PLANT has CONSTRAINTFACTOR = 0 in the Source Constraint extract."""
        miss_part = self._missing_cols(part_df, [PROCUREMENT_COL, PART_MATERIAL_COL, PART_PLANT_COL])
        miss_sc   = self._missing_cols(sc_df, [SC_MATERIAL_COL, SC_PLANT_COL, SC_FACTOR_COL])
        if miss_part or miss_sc:
            print(f"⚠️   Sub-rule (b) skipped — missing column(s): Part={miss_part}, SourceConstraint={miss_sc}")
            return {}

        sc = pd.DataFrame({
            "mat":    self._clean(sc_df[SC_MATERIAL_COL]),
            "plant":  self._clean(sc_df[SC_PLANT_COL]),
            "factor": self._clean(sc_df[SC_FACTOR_COL]),
        })

        # Per combination: list of raw factors, and whether any of them is zero
        factors_by_key = {}
        zero_keys      = set()
        for mat, plant, factor in zip(sc["mat"], sc["plant"], sc["factor"]):
            key = (mat, plant)
            factors_by_key.setdefault(key, []).append(factor)
            num = self._parse_number(factor)
            if num is not None and num == 0:
                zero_keys.add(key)

        rows   = self._proc_rows(part_df, PROCUREMENT_VALUE_E)
        result = {}
        for idx, mat, plant in zip(rows.index, rows["mat"], rows["plant"]):
            key = (mat, plant)
            if key in zero_keys:
                result[idx] = (
                    f"PROCUREMENTTYPE: 'E' but {SC_FACTOR_COL} is '0' for {SC_MATERIAL_COL} "
                    f"'{mat}' + {SC_PLANT_COL} '{plant}' in Source Constraint extract"
                )
            elif MISSING_IN_CONSTRAINT_IS_ERROR and key not in factors_by_key:
                result[idx] = (
                    f"PROCUREMENTTYPE: 'E' but {SC_MATERIAL_COL} '{mat}' + "
                    f"{SC_PLANT_COL} '{plant}' is not present in Source Constraint extract"
                )
        return result

    # ── Numeric field rule (>= 0) ──
    def check_numeric_field(self, part_df: pd.DataFrame, field: str, mode: str) -> dict:
        """
        mode '>0'  : value must be greater than 0
        mode '>=0' : value must be greater than or equal to 0
        Blank -> skipped unless BLANK_FAILS_NUMERIC_CHECKS; non-numeric -> always flagged.
        Returns {row_idx: reason}.
        """
        if field not in part_df.columns:
            print(f"⚠️   Rule '{field}' skipped — column not found in Part extract")
            return {}

        raw   = self._clean(part_df[field])
        blank = raw == ""
        num   = pd.to_numeric(raw, errors="coerce")
        non_numeric = (~blank) & num.isna()

        if mode == ">0":
            bad_value = num.notna() & (num <= 0)
        else:  # ">=0"
            bad_value = num.notna() & (num < 0)

        fail = bad_value | non_numeric
        if BLANK_FAILS_NUMERIC_CHECKS:
            fail = fail | blank

        reason = FIELD_REASON[field]
        return {idx: reason for idx in part_df.index[fail]}

    # ── NEW: MINREMSHELFLIFE > 0 when BATCHMANAGED = 'Y' ──
    def check_minshelflife_batchmanaged(self, part_df: pd.DataFrame) -> dict:
        miss = self._missing_cols(part_df, [MINSHELFLIFE_COL, BATCHMANAGED_COL])
        if miss:
            print(f"⚠️   Rule '{MINSHELFLIFE_COL}' skipped — column(s) not found in Part extract: {miss}")
            return {}

        batch = self._clean(part_df[BATCHMANAGED_COL]).str.upper()
        in_scope = batch == BATCHMANAGED_YES.upper()

        raw   = self._clean(part_df[MINSHELFLIFE_COL])
        blank = raw == ""
        num   = pd.to_numeric(raw, errors="coerce")
        non_numeric = (~blank) & num.isna()
        not_positive = num.notna() & (num <= 0)

        fail = in_scope & (blank | non_numeric | not_positive)

        result = {}
        for idx in part_df.index[fail]:
            if blank.loc[idx]:
                result[idx] = (f"MINREMSHELFLIFE: is blank but {BATCHMANAGED_COL} is "
                               f"'{BATCHMANAGED_YES}' (must be greater than 0)")
            elif non_numeric.loc[idx]:
                result[idx] = (f"MINREMSHELFLIFE: '{raw.loc[idx]}' is not a valid number "
                               f"while {BATCHMANAGED_COL} is '{BATCHMANAGED_YES}'")
            else:
                result[idx] = (f"MINREMSHELFLIFE: '{raw.loc[idx]}' must be greater than 0 "
                               f"when {BATCHMANAGED_COL} is '{BATCHMANAGED_YES}'")
        return result

    # ── STANDARDPRICE sub-rule (a): non-zero for non FERT / HAWA rows ──
    def check_standard_price(self, part_df: pd.DataFrame, type_col):
        if STANDARDPRICE_COL not in part_df.columns:
            print(f"⚠️   Rule '{STANDARDPRICE_COL}' skipped — column not found in Part extract")
            return {}
        if type_col is None:
            print(f"⚠️   Rule '{STANDARDPRICE_COL}' skipped — none of "
                  f"{PRODUCT_TYPE_COL_CANDIDATES} found in Part extract")
            return {}

        ptype    = self._clean(part_df[type_col]).str.upper()
        in_scope = ~ptype.isin(FG_HAWA_TYPES)          # blank type counts as non FERT/HAWA

        raw   = self._clean(part_df[STANDARDPRICE_COL])
        blank = raw == ""
        num   = pd.to_numeric(raw, errors="coerce")
        zero        = num.notna() & (num == 0)
        non_numeric = (~blank) & num.isna()

        fail = in_scope & (zero | non_numeric)
        if BLANK_FAILS_NUMERIC_CHECKS:
            fail = fail | (in_scope & blank)

        result = {}
        for idx in part_df.index[fail]:
            t = ptype.loc[idx] or "(blank)"
            if zero.loc[idx]:
                result[idx] = f"STANDARDPRICE: is 0 for a non FERT/HAWA row ({type_col} '{t}')"
            elif blank.loc[idx]:
                result[idx] = f"STANDARDPRICE: is blank for a non FERT/HAWA row ({type_col} '{t}')"
            else:
                result[idx] = (f"STANDARDPRICE: '{raw.loc[idx]}' is not a valid number "
                               f"for a non FERT/HAWA row ({type_col} '{t}')")
        return result

    # ── NEW: Part class derivation (vectorised) ──
    def derive_part_class(self, part_df: pd.DataFrame, type_col) -> tuple:
        """
        Returns (part_class, ibp, mrp) as Series indexed like part_df.
        part_class is 'SFG' / 'RM' / 'PM' or '' (no class).
        Missing IBPSTATUS / MRPTYPE columns are treated as blank (a warning is printed).
        """
        ptype = self._clean(part_df[type_col]).str.upper()

        if IBPSTATUS_COL in part_df.columns:
            ibp = self._clean(part_df[IBPSTATUS_COL]).str.upper()
        else:
            print(f"⚠️   {IBPSTATUS_COL} column not found in Part extract — treated as blank "
                  f"for part class derivation (HAWA rows will not get a class; HIBE rows will be PM)")
            ibp = pd.Series("", index=part_df.index)

        if MRPTYPE_COL in part_df.columns:
            mrp = self._clean(part_df[MRPTYPE_COL]).str.upper()
        else:
            print(f"⚠️   {MRPTYPE_COL} column not found in Part extract — treated as blank "
                  f"for part class derivation (RAWT rows will not get a class)")
            mrp = pd.Series("", index=part_df.index)

        ibp_x0 = ibp.isin({"X0", "XO"})
        mrp_x0 = mrp.isin({"X0", "XO"})

        sfg = ptype.isin({"BLND", "HALB", "BASB"}) | ((ptype == "RAWT") & mrp_x0)
        rm  = (ptype == "ROH")  | ((ptype == "HAWA") & (ibp == "N"))
        pm  = (ptype == "VERP") | ((ptype == "HIBE") & ((ibp == "") | ibp_x0))

        part_class = pd.Series("", index=part_df.index)
        part_class[sfg] = "SFG"
        part_class[rm]  = "RM"
        part_class[pm]  = "PM"
        return part_class, ibp, mrp

    # ── NEW: STANDARDPRICE sub-rule (b): non-zero for RM / PM / SFG part class ──
    def check_standard_price_partclass(self, part_df: pd.DataFrame, type_col) -> dict:
        if STANDARDPRICE_COL not in part_df.columns:
            print(f"⚠️   Rule '{STANDARDPRICE_COL}' (b) skipped — column not found in Part extract")
            return {}
        if type_col is None:
            print(f"⚠️   Rule '{STANDARDPRICE_COL}' (b) skipped — none of "
                  f"{PRODUCT_TYPE_COL_CANDIDATES} found in Part extract")
            return {}

        part_class, ibp, mrp = self.derive_part_class(part_df, type_col)
        ptype    = self._clean(part_df[type_col]).str.upper()
        in_scope = part_class.isin(PART_CLASS_SCOPE)

        raw   = self._clean(part_df[STANDARDPRICE_COL])
        blank = raw == ""
        num   = pd.to_numeric(raw, errors="coerce")
        zero        = num.notna() & (num == 0)
        non_numeric = (~blank) & num.isna()

        # Blank always fails here: a blank price cannot be non-zero
        fail = in_scope & (zero | non_numeric | blank)

        result = {}
        for idx in part_df.index[fail]:
            ctx = (f"part class {part_class.loc[idx]}: {type_col} '{ptype.loc[idx]}', "
                   f"{IBPSTATUS_COL} '{ibp.loc[idx]}', {MRPTYPE_COL} '{mrp.loc[idx]}'")
            if zero.loc[idx]:
                result[idx] = f"STANDARDPRICE: is 0 for a row with {ctx}"
            elif blank.loc[idx]:
                result[idx] = f"STANDARDPRICE: is blank for a row with {ctx}"
            else:
                result[idx] = f"STANDARDPRICE: '{raw.loc[idx]}' is not a valid number for a row with {ctx}"
        return result


# ══════════════════════════════════════════════
#  Validator
# ══════════════════════════════════════════════
class PartBusinessTableValidator:

    def __init__(self, part_path: str, ps_path: str, pb_path: str, sc_path: str):
        self.part_path = part_path
        self.ps_path   = ps_path
        self.pb_path   = pb_path
        self.sc_path   = sc_path
        self.df        = pd.DataFrame()
        self.ps_df     = pd.DataFrame()
        self.pb_df     = pd.DataFrame()
        self.sc_df     = pd.DataFrame()
        self.product_type_col = None
        self.error_map  = {}   # row_idx -> [failed field names]
        self.reason_map = {}   # row_idx -> {field: reason}
        # field -> {row_idx: {sub_rule_id: reason}}   (PROCUREMENTTYPE and STANDARDPRICE)
        self.sub_map    = {f: {} for f in FIELDS_WITH_SUBS}

    @staticmethod
    def _read(path: str) -> pd.DataFrame:
        df = pd.read_csv(path, sep="\t", dtype=str)
        df.columns = [c.strip().upper() for c in df.columns]
        return df

    def load(self):
        self.df    = self._read(self.part_path)
        self.ps_df = self._read(self.ps_path)
        self.pb_df = self._read(self.pb_path)
        self.sc_df = self._read(self.sc_path)

        self.product_type_col = next(
            (c for c in PRODUCT_TYPE_COL_CANDIDATES if c in self.df.columns), None
        )

    def _add_error(self, idx, field: str, reason: str):
        self.error_map.setdefault(idx, [])
        if field not in self.error_map[idx]:
            self.error_map[idx].append(field)
        self.reason_map.setdefault(idx, {})[field] = reason

    def _register_subs(self, field: str, sub_results: list):
        """sub_results: [(sub_id, {row_idx: reason}), ...] -> fills sub_map and error maps."""
        for sub_id, reasons in sub_results:
            for idx, reason in reasons.items():
                self.sub_map[field].setdefault(idx, {})[sub_id] = reason

        for idx, subs in self.sub_map[field].items():
            self._add_error(
                idx, field,
                " | ".join(f"[{sid}] {txt}" for sid, txt in subs.items()),
            )

    def validate(self):
        engine = PartBusinessRuleEngine()

        # ── PROCUREMENTTYPE (sub-rules a, b, c) ──
        self._register_subs(RULE_FIELD, [
            (SUB_A, engine.check_partsource(self.df, self.ps_df)),
            (SUB_B, engine.check_constraint_factor(self.df, self.sc_df)),
            (SUB_C, engine.check_partsource_buy(self.df, self.pb_df)),
        ])

        # ── MINREMSHELFLIFE  (> 0 when BATCHMANAGED = 'Y') ──
        for idx, reason in engine.check_minshelflife_batchmanaged(self.df).items():
            self._add_error(idx, MINSHELFLIFE_COL, reason)

        # ── INTERVALSTOEXPIRYDAYS  (>= 0) ──
        for idx, reason in engine.check_numeric_field(self.df, INTERVALSEXPIRY_COL, ">=0").items():
            self._add_error(idx, INTERVALSEXPIRY_COL, reason)

        # ── STANDARDPRICE (a: non FERT/HAWA, b: RM/PM/SFG part class) ──
        self._register_subs(STANDARDPRICE_COL, [
            (SUB_A, engine.check_standard_price(self.df, self.product_type_col)),
            (SUB_B, engine.check_standard_price_partclass(self.df, self.product_type_col)),
        ])

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

    def get_sub_rule_counts(self, field: str = RULE_FIELD) -> dict:
        sub_ids = FIELDS_WITH_SUBS[field][0]
        counts  = {sid: 0 for sid in sub_ids}
        for subs in self.sub_map[field].values():
            for sid in subs:
                counts[sid] += 1
        return counts


# ══════════════════════════════════════════════
#  Report Writer
# ══════════════════════════════════════════════
class PartBusinessReportWriter:

    SHEET_SUMMARY = "Summary"
    SHEET_RULES   = "Rules"

    RULES_CONTENT = {
        RULE_FIELD: [
            "a) If PROCUREMENTTYPE is 'E' for a MATERIALNUMBER-PLANT combination in the "
            "Part extract, that combination must be present in the PartSource (Make) "
            "extract (MATERIAL and PLANT columns).",
            "b) If PROCUREMENTTYPE is 'E' for a MATERIALNUMBER-PLANT combination in the "
            "Part extract, CONSTRAINTFACTOR must not be '0' for that MATERIAL-PLANT "
            "combination in the Source Constraint extract.",
            "c) If PROCUREMENTTYPE is 'F' for a MATERIALNUMBER-PLANT combination in the "
            "Part extract, that combination must be present in the PartSource (Buy) "
            "extract (MATERIALNUMBER and PLANT columns).",
        ],
        MINSHELFLIFE_COL: [
            "When BATCHMANAGED = 'Y', MINREMSHELFLIFE must be a number greater than 0 "
            "(blank, 0, negative or non-numeric values are flagged).",
        ],
        INTERVALSEXPIRY_COL: [
            "INTERVALSTOEXPIRYDAYS must be a number greater than or equal to 0.",
        ],
        STANDARDPRICE_COL: [
            "a) STANDARDPRICE must be non-zero for rows where the product type "
            "(PRODUCTYPE column) is not FERT or HAWA (e.g. HALB, ROH). "
            "FERT and HAWA rows are not checked by this sub-rule.",
            "b) STANDARDPRICE must be non-zero (blank, 0 or non-numeric values are flagged) "
            "for rows whose part class is RM, PM or SFG. Part class is derived as: "
            "BLND / HALB / BASB = SFG; HAWA with IBPSTATUS 'N' = RM; ROH = RM; "
            "HIBE with IBPSTATUS blank or 'X0' = PM; VERP = PM; "
            "RAWT with MRPTYPE 'X0' = SFG. Other combinations have no part class "
            "and are not checked by this sub-rule.",
        ],
    }

    def __init__(self, validator: PartBusinessTableValidator, output_path: str):
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

    def _highlight_cols_for(self, field_name: str, row_idx) -> list:
        """Part columns highlighted red for this row when the given rule fails."""
        v = self.validator
        if field_name == RULE_FIELD:
            return [PROCUREMENT_COL, PART_MATERIAL_COL, PART_PLANT_COL]

        if field_name == STANDARDPRICE_COL:
            cols = [STANDARDPRICE_COL]
            if v.product_type_col:
                cols.append(v.product_type_col)               # cross-field rule: both involved
            # Sub-rule (b) also depends on IBPSTATUS and MRPTYPE
            if SUB_B in v.sub_map[STANDARDPRICE_COL].get(row_idx, {}):
                cols.extend([IBPSTATUS_COL, MRPTYPE_COL])
            return cols

        if field_name == MINSHELFLIFE_COL:
            return [MINSHELFLIFE_COL, BATCHMANAGED_COL]       # cross-field rule: both involved

        return [field_name]

    # ══════════════════════════════════════════
    #  Summary sheet  (field row + sub-rows for fields with sub-rules)
    # ══════════════════════════════════════════
    def _write_summary_sheet_into(self, ws, error_map: dict, total_rows: int):

        # ── Row 1 : Title ──
        ws.merge_cells("A1:G1")
        title_cell           = ws.cell(row=1, column=1, value="Part Business Validation Summary")
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

        col_error_counts: dict = {}
        for bad_cols in error_map.values():
            for col in bad_cols:
                col_error_counts[col] = col_error_counts.get(col, 0) + 1

        row_num   = 3
        field_num = 1

        for col_name in FIELD_ORDER:
            count      = col_error_counts.get(col_name, 0)
            pct_error  = round((count / total_rows) * 100, 2) if total_rows else 0
            pct_health = round(100 - pct_error, 2)

            has_subs = col_name in FIELDS_WITH_SUBS
            # Fields with sub-rows show their reasons on the sub-rows instead
            reason   = "" if has_subs else (FIELD_REASON.get(col_name, "") if count > 0 else "")

            ws.cell(row=row_num, column=1, value=field_num)
            ws.cell(row=row_num, column=2, value=col_name)
            ws.cell(row=row_num, column=3, value=count)
            ws.cell(row=row_num, column=4, value=total_rows)
            ws.cell(row=row_num, column=5, value=f"{pct_health}%")
            ws.cell(row=row_num, column=6, value=f"{pct_error}%")
            ws.cell(row=row_num, column=7, value=reason)
            self._style_summary_data_row(ws, row_num, bold=True, fill=WHITE_FILL)
            ws.cell(row=row_num, column=7).alignment = Alignment(
                horizontal="left", vertical="center", wrap_text=True
            )
            row_num += 1

            # Sub-rule rows (PROCUREMENTTYPE and STANDARDPRICE)
            if has_subs:
                sub_ids, sub_reason_map = FIELDS_WITH_SUBS[col_name]
                sub_counts = self.validator.get_sub_rule_counts(col_name)

                for sid in sub_ids:
                    s_count = sub_counts.get(sid, 0)
                    s_err   = round((s_count / total_rows) * 100, 2) if total_rows else 0
                    s_hlth  = round(100 - s_err, 2)

                    ws.cell(row=row_num, column=1, value="")
                    ws.cell(row=row_num, column=2, value=f"   {col_name} - sub-rule {sid})")
                    ws.cell(row=row_num, column=3, value=s_count)
                    ws.cell(row=row_num, column=4, value=total_rows)
                    ws.cell(row=row_num, column=5, value=f"{s_hlth}%")
                    ws.cell(row=row_num, column=6, value=f"{s_err}%")
                    ws.cell(row=row_num, column=7,
                            value=sub_reason_map[sid] if s_count > 0 else "")
                    self._style_summary_data_row(ws, row_num, fill=WHITE_FILL, italic=True)
                    ws.cell(row=row_num, column=2).alignment = Alignment(horizontal="left", vertical="center")
                    ws.cell(row=row_num, column=7).alignment = Alignment(
                        horizontal="left", vertical="center", wrap_text=True
                    )
                    row_num += 1

            field_num += 1

        # ── TOTAL row (sum of per-field failing rows) ──
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

        col_widths = [6, 42, 14, 16, 12, 12, 90]
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

            # Sort so rows of the same MATERIALNUMBER-PLANT sit together
            sort_cols = [c for c in (PART_MATERIAL_COL, PART_PLANT_COL) if c in subset.columns]
            if sort_cols:
                subset = subset.sort_values(sort_cols, kind="stable", na_position="last")

            self._write_header(ws, subset.columns)
            col_idx_map = {col: i for i, col in enumerate(subset.columns, start=1)}

            for excel_row, (orig_idx, row_data) in enumerate(subset.iterrows(), start=2):
                for c_idx, (col, value) in enumerate(zip(subset.columns, row_data), start=1):
                    cell           = ws.cell(row=excel_row, column=c_idx, value=value)
                    cell.font      = BODY_FONT
                    cell.alignment = Alignment(vertical="center")
                    cell.fill      = WHITE_FILL
                    cell.border    = THIN_BORDER

                # Only the columns involved in THIS row's failing rule are highlighted
                for target_col in self._highlight_cols_for(field_name, orig_idx):
                    if target_col in col_idx_map:
                        target_cell      = ws.cell(row=excel_row, column=col_idx_map[target_col])
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
        title_cell           = ws.cell(row=1, column=1, value="Part Table – Business Validation Rules")
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
        ws.column_dimensions["B"].width = 30
        ws.column_dimensions["C"].width = 100

    # ── Main write ────────────────────────────
    def write(self):
        v  = self.validator
        df = v.df.copy()

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

        proc_counts        = v.get_sub_rule_counts(RULE_FIELD)
        price_counts       = v.get_sub_rule_counts(STANDARDPRICE_COL)
        errors_by_field    = v.get_errors_by_field()
        fields_with_errors = [f for f in FIELD_ORDER if f in errors_by_field]
        print(f"\n✅  Output saved  → {self.output_path}")
        print(f"   Total rows    : {len(df)}")
        print(f"   Error rows    : {len(v.error_map)}")
        print(f"   PROCUREMENTTYPE sub-rule (a) : {proc_counts[SUB_A]}")
        print(f"   PROCUREMENTTYPE sub-rule (b) : {proc_counts[SUB_B]}")
        print(f"   PROCUREMENTTYPE sub-rule (c) : {proc_counts[SUB_C]}")
        print(f"   STANDARDPRICE   sub-rule (a) : {price_counts[SUB_A]}")
        print(f"   STANDARDPRICE   sub-rule (b) : {price_counts[SUB_B]}")
        for f in (MINSHELFLIFE_COL, INTERVALSEXPIRY_COL):
            print(f"   {f:28s}: {len(errors_by_field.get(f, []))}")
        print(f"   Field sheets  : {fields_with_errors}")


# ══════════════════════════════════════════════
#  Orchestrator
# ══════════════════════════════════════════════
class PartBusinessTableProcessor:

    def __init__(self, part_path: str, ps_path: str, pb_path: str, sc_path: str, output_path: str):
        self.validator = PartBusinessTableValidator(part_path, ps_path, pb_path, sc_path)
        self.writer    = PartBusinessReportWriter(self.validator, output_path)

    def run(self):
        print("📂  Loading files …")
        self.validator.load()
        print(f"    Part columns detected             : {list(self.validator.df.columns)}")
        print(f"    PartSource (Make) columns detected: {list(self.validator.ps_df.columns)}")
        print(f"    PartSource (Buy) columns detected : {list(self.validator.pb_df.columns)}")
        print(f"    Source Constraint columns detected: {list(self.validator.sc_df.columns)}")
        print(f"    Product type column used          : {self.validator.product_type_col}")
        print("🔍  Validating business rules …")
        self.validator.validate()
        print("📝  Writing report …")
        self.writer.write()


# ══════════════════════════════════════════════
#  Entry Point
# ══════════════════════════════════════════════
if __name__ == "__main__":
    processor = PartBusinessTableProcessor(
        part_path   = PART_INPUT_FILE,
        ps_path     = PARTSOURCE_MAKE_INPUT_FILE,
        pb_path     = PARTSOURCE_BUY_INPUT_FILE,
        sc_path     = SOURCE_CONSTRAINT_INPUT_FILE,
        output_path = OUTPUT_FILE,
    )
    processor.run()
