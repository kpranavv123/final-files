

"""
PRODUCTHIERARCHY (FG) - Fast Excel Data Validation Tool
Requires: pandas, numpy, xlsxwriter
"""
import os
from collections import defaultdict

import numpy as np
import pandas as pd
import xlsxwriter

INPUT_FILE  = r"C:\Users\SW526XH\Downloads\Go Live-1\ProductH\ProductHierarchy_Without_Z3_Materials.tab"
OUTPUT_FILE = r"C:\Users\SW526XH\Downloads\Go Live-1\ProductH\Validated_Product Hierarchy_Technical.xlsx"

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
    "GROSSWEIGHT": ["Field must be greater than 0.", "GROSSWEIGHT must be >= NETWEIGHT."],
    "MATERIALTYPE": ["Must not be blank."],
    "IBPSTATUS": ["Allowed values: IBP, IBP EXP or blank.", "Any other value is treated as an error."],
    "DUPLICATE_CHECK": ["The MATERIALNUMBER column must not contain duplicate values across the extract."],
}

FIELD_REASON_MAP = {
    **{f: f"{f}: is blank." for f in NOT_BLANK_FIELDS},
    **{f: f"{f}: must be greater than 0." for f in POSITIVE_FIELDS},
    "MATERIALTYPE": "MATERIALTYPE: is blank.",
    "IBPSTATUS": "IBPSTATUS: Invalid value - must be 'IBP', 'IBP EXP' or blank.",
    "DUPLICATE_CHECK": DUPLICATE_SUMMARY_LABEL,
}

SUB_ROWS = {  # field -> [(label, mask_key, reason)]
    "GROSSWEIGHT": [
        ("  ↳ Not greater than 0", "GW_POS", FIELD_REASON_MAP["GROSSWEIGHT"]),
        ("  ↳ Less than NETWEIGHT", "GW_NET", GW_VS_NET_REASON),
    ]
}


# ───────────────────────── VALIDATION (vectorized) ─────────────────────────
def is_blank(s: pd.Series) -> pd.Series:
    t = s.fillna("").astype(str).str.strip()
    return (t == "") | (t.str.lower() == "nan")


def validate(df: pd.DataFrame):
    df = df.copy()
    df.columns = df.columns.str.strip().str.upper()
    for c in ["MATERIALTYPE", "IBPSTATUS", "MATERIALNUMBER", "NETWEIGHT", "GROSSWEIGHT"]:
        if c not in df.columns:
            df[c] = ""
    df["MATERIALTYPE"] = df["MATERIALTYPE"].fillna("").astype(str).str.strip().str.upper()

    masks = {}      # mask_key -> numpy bool array
    ordered = []    # (mask_key, reason) in the order reasons should appear

    def add(key, mask, reason):
        masks[key] = np.asarray(mask, dtype=bool)
        ordered.append((key, reason))

    for f in NOT_BLANK_FIELDS:
        if f in df.columns:
            add(f, is_blank(df[f]), FIELD_REASON_MAP[f])

    nums = {}
    for f in POSITIVE_FIELDS:
        if f in df.columns:
            nums[f] = pd.to_numeric(df[f], errors="coerce")
            m = nums[f].isna() | (nums[f] <= 0)
            add("GW_POS" if f == "GROSSWEIGHT" else f, m, FIELD_REASON_MAP[f])

    add("MATERIALTYPE", is_blank(df["MATERIALTYPE"]), FIELD_REASON_MAP["MATERIALTYPE"])

    ibp_blank = is_blank(df["IBPSTATUS"])
    ibp_up = df["IBPSTATUS"].fillna("").astype(str).str.strip().str.upper()
    add("IBPSTATUS", (~ibp_blank) & (~ibp_up.isin(VALID_IBP_STATUSES)), FIELD_REASON_MAP["IBPSTATUS"])

    g, n = nums.get("GROSSWEIGHT"), nums.get("NETWEIGHT")
    add("GW_NET", g.notna() & n.notna() & (g < n), GW_VS_NET_REASON)

    mat = df["MATERIALNUMBER"].fillna("").astype(str).str.strip()
    add("DUPLICATE_CHECK", (mat != "") & mat.duplicated(keep=False), DUPLICATE_SUMMARY_LABEL)

    # Build ERROR_COLUMNS with numpy string concatenation (no row loops)
    err = np.full(len(df), "", dtype=object)
    for key, reason in ordered:
        err = err + np.where(masks[key], reason + " | ", "")
    df["ERROR_COLUMNS"] = pd.Series(err, index=df.index).str.slice(0, -3)

    # Field-level masks used by the summary (GROSSWEIGHT = union of its two checks)
    field_masks = {k: v for k, v in masks.items() if k not in ("GW_POS", "GW_NET")}
    field_masks["GROSSWEIGHT"] = masks["GW_POS"] | masks["GW_NET"]

    print(f"    Duplicate MATERIALNUMBER rows: {masks['DUPLICATE_CHECK'].sum():,}")
    return df, masks, field_masks


# ───────────────────────── EXCEL WRITER ─────────────────────────
def build_report(df, masks, field_masks, out_path):
    wb = xlsxwriter.Workbook(out_path, {"nan_inf_to_errors": True})

    def F(**kw):
        base = {"font_name": "Arial", "font_size": 10}
        base.update(kw)
        return wb.add_format(base)

    border = {"border": 1}
    hdr_f   = F(bold=True, bg_color="#D9E1F2", align="center", valign="vcenter", text_wrap=True, **border)
    title_f = F(bold=True, font_size=14, bg_color="#BDD7EE", align="center", valign="vcenter")
    title2_f= F(bold=True, font_size=13, bg_color="#BDD7EE", align="center", valign="vcenter")
    sh_f    = F(bold=True, bg_color="#BDD7EE", align="center", valign="vcenter", text_wrap=True, **border)
    body_c  = F(align="center", valign="vcenter", text_wrap=True, **border)
    body_l  = F(align="left", valign="vcenter", text_wrap=True, **border)
    sub_lbl = F(italic=True, align="left", valign="vcenter", indent=1, **border)
    sub_c   = F(italic=True, align="center", valign="vcenter", **border)
    sub_l   = F(italic=True, align="left", valign="vcenter", text_wrap=True, **border)
    tot_f   = F(bold=True, bg_color="#F2F2F2", align="center", valign="vcenter", **border)
    stat_l  = F(bold=True, bg_color="#EDEDED", align="left", valign="vcenter", **border)
    rule_c  = F(bg_color="#E2EFDA", align="center", valign="vcenter", text_wrap=True, **border)
    rule_cb = F(bold=True, bg_color="#E2EFDA", align="center", valign="vcenter", text_wrap=True, **border)
    row_err = F(bg_color="#FFF2CC", align="center", valign="vcenter", **border)
    cell_err= F(bg_color="#FF0000", font_color="#FFFFFF", bold=True, align="center", valign="vcenter", **border)

    display = df.fillna("")
    headers = list(display.columns)
    col_idx = {h: i for i, h in enumerate(headers)}
    n_rows, n_cols = len(display), len(headers)
    has_err = (display["ERROR_COLUMNS"] != "").values

    def widths(ws, d, sample=2000, mn=10, mx=45):
        s = d.head(sample)
        for i, c in enumerate(d.columns):
            try:
                dl = int(s[c].astype(str).str.len().max())
            except Exception:
                dl = 0
            ws.set_column(i, i, min(max(max(len(str(c)), dl) + 2, mn), mx))

    # ── Sheet 1: main (colored) ──
    print("    Writing main sheet …")
    ws = wb.add_worksheet("PRODUCTHIERARCHY_FG")
    ws.write_row(0, 0, headers, hdr_f)
    ws.set_row(0, 30)
    ws.freeze_panes(1, 0)

    values = display.values.tolist()
    for r, row in enumerate(values, start=1):
        if has_err[r - 1]:
            ws.write_row(r, 0, row, row_err)
        else:
            ws.write_row(r, 0, row)

    # red cells: only where errors exist
    red_targets = {k: k for k in field_masks}
    red_targets["DUPLICATE_CHECK"] = "MATERIALNUMBER"
    for key, mask in list(masks.items()) + [("GROSSWEIGHT", field_masks["GROSSWEIGHT"])]:
        if key == "GW_POS":
            col = "GROSSWEIGHT"
        elif key == "GW_NET":
            col = None  # handled below (GROSSWEIGHT + NETWEIGHT)
        elif key == "DUPLICATE_CHECK":
            col = "MATERIALNUMBER"
        else:
            col = key
        if key == "GW_NET":
            for r in np.flatnonzero(mask):
                for c in ("GROSSWEIGHT", "NETWEIGHT"):
                    ws.write(r + 1, col_idx[c], values[r][col_idx[c]], cell_err)
            continue
        if col not in col_idx or key == "GROSSWEIGHT":
            continue
        ci = col_idx[col]
        for r in np.flatnonzero(mask):
            ws.write(r + 1, ci, values[r][ci], cell_err)
    widths(ws, display)

    # ── Sheet 2: Summary ──
    print("    Writing Summary …")
    ws = wb.add_worksheet("Summary")
    ws.merge_range(0, 0, 0, 6, "ProductHierarchy FG Validation Summary", title_f)
    ws.set_row(0, 24)
    ws.write_row(1, 0, ["#", "Field Name", "Error Count", "Record Count", "% Health", "% of Error", "Reason"], sh_f)

    total = n_rows
    row = 2
    total_errors = 0
    for num, name in enumerate(ALL_SUMMARY_FIELDS, start=1):
        cnt = int(field_masks[name].sum()) if name in field_masks else 0
        total_errors += cnt
        pe = round(cnt / total * 100, 2) if total else 0
        reason = "" if (cnt == 0 or name in SUB_ROWS) else FIELD_REASON_MAP.get(name, f"{name}: is blank.")
        ws.write_row(row, 0, [num, name, cnt, total, f"{round(100 - pe, 2)}%", f"{pe}%"], body_c)
        ws.write(row, 6, reason, body_l)
        row += 1
        if name in SUB_ROWS and cnt > 0:
            for label, key, sreason in SUB_ROWS[name]:
                sc = int(masks[key].sum())
                spe = round(sc / total * 100, 2) if total else 0
                ws.write(row, 0, "", sub_c)
                ws.write(row, 1, label, sub_lbl)
                ws.write_row(row, 2, [sc, total, f"{round(100 - spe, 2)}%", f"{spe}%"], sub_c)
                ws.write(row, 6, sreason, sub_l)
                row += 1

    trc = total * len(ALL_SUMMARY_FIELDS)
    tpe = round(total_errors / trc * 100, 2) if trc else 0
    ws.write_row(row, 0, ["", "TOTAL", total_errors, trc, f"{round(100 - tpe, 2)}%", f"{tpe}%", ""], tot_f)
    row += 2
    rec_err = int(has_err.sum())
    for label, val in [("Total Records:", total), ("Records with Errors:", rec_err),
                       ("Records Passing:", total - rec_err)]:
        ws.merge_range(row, 0, row, 1, label, stat_l)
        ws.write(row, 2, val, body_c)
        row += 1
    for i, w in enumerate([6, 28, 12, 13, 11, 12, 55]):
        ws.set_column(i, i, w)

    # ── Sheet 3: Rule_Set ──
    print("    Writing Rule_Set …")
    ws = wb.add_worksheet("Rule_Set")
    ws.merge_range(0, 0, 0, 2, "ProductHierarchy FG – Validation Rules", title2_f)
    ws.set_row(0, 22)
    ws.write_row(2, 0, ["#", "Field", "Rule Description"], hdr_f)
    row, rn = 3, 1
    for field, rules in RULES_CONTENT.items():
        if len(rules) > 1:
            ws.merge_range(row, 0, row + len(rules) - 1, 0, rn, rule_cb)
            ws.merge_range(row, 1, row + len(rules) - 1, 1, field, rule_cb)
        else:
            ws.write(row, 0, rn, rule_cb)
            ws.write(row, 1, field, rule_cb)
        for i, txt in enumerate(rules):
            ws.write(row + i, 2, txt, body_c)
        row += len(rules)
        rn += 1
    ws.set_column(0, 0, 6); ws.set_column(1, 1, 30); ws.set_column(2, 2, 65)

    # ── Per-field error sheets (NO colors → fast) ──
    print("    Writing per-field error sheets …")
    plain_hdr = F(bold=True)
    sheet_fields = {k: m for k, m in field_masks.items() if m.any()}
    used = {"PRODUCTHIERARCHY_FG", "Summary", "Rule_Set"}
    for field in sorted(sheet_fields):
        m = sheet_fields[field]
        subset = display.loc[m].copy()
        if field == "GROSSWEIGHT":
            p, q = masks["GW_POS"][m], masks["GW_NET"][m]
            subset["ERROR_COLUMNS"] = np.where(
                p & q, FIELD_REASON_MAP["GROSSWEIGHT"] + " | " + GW_VS_NET_REASON,
                np.where(p, FIELD_REASON_MAP["GROSSWEIGHT"], GW_VS_NET_REASON))
        else:
            subset["ERROR_COLUMNS"] = FIELD_REASON_MAP.get(field, f"{field}: is blank.")

        name = field[:31]
        k = 1
        while name in used:
            name = f"{field[:28]}_{k}"; k += 1
        used.add(name)

        ws = wb.add_worksheet(name)
        ws.write_row(0, 0, list(subset.columns), plain_hdr)
        ws.freeze_panes(1, 0)
        for r, rv in enumerate(subset.values.tolist(), start=1):
            ws.write_row(r, 0, rv)
        ws.write(len(subset) + 2, 0, f"Total error rows for '{field}': {len(subset)}")
        widths(ws, subset)

    wb.close()
    print(f"\n✅  Report saved → {out_path}")


# ───────────────────────── PIPELINE ─────────────────────────
def read_input(path):
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xlsm", ".xls"):
        return pd.read_excel(path, dtype=str)
    if ext in (".csv", ".tab"):
        return pd.read_csv(path, dtype=str, sep="\t", encoding="latin-1", low_memory=False)
    raise ValueError(f"Unsupported file type: {ext}")


if __name__ == "__main__":
    print(f"📂  Reading: {INPUT_FILE}")
    data = read_input(INPUT_FILE)
    print(f"🔍  Validating {len(data):,} rows …")
    validated, masks, field_masks = validate(data)
    print(f"⚠️   Errors found in {(validated['ERROR_COLUMNS'] != '').sum():,} row(s)")
    print("📝  Building Excel report …")
    build_report(validated, masks, field_masks, OUTPUT_FILE)
