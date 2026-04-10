import pandas as pd
import numpy as np

# ── Custom Exceptions ────────────────────────────────────────────────────────

class WorkbookValidationError(Exception):
    """Raised when an uploaded workbook fails initial structure validation."""
    def __init__(self, validation_results):
        self.validation_results = validation_results
        super().__init__("Workbook validation failed.")

class SheetStructureError(Exception):
    """Raised for missing sheets or missing columns."""
    pass

class SalesBlockParseError(Exception):
    """Raised when the sales data block layout does not match expectations."""
    pass

# ── Constants ────────────────────────────────────────────────────────
NEEDED_MATURITIES = [85, 95, 105, 115]
YEARS_NEEDED = list(range(2, 11))

def resolve_sheet_name(xls: pd.ExcelFile, keywords: list) -> str:
    """Finds the first sheet name containing any of the keywords (case-insensitive)."""
    for sheet in xls.sheet_names:
        low_sheet = sheet.lower()
        if any(k.lower() in low_sheet for k in keywords):
            return sheet
    return None


# ── Helper Functions ─────────────────────────────────────────────────────────

def _norm_txt(x):
    if pd.isna(x):
        return ""
    return str(x).strip().lower()

def _find_cell(sheet_df, text):
    target = text.strip().lower()
    for r in range(sheet_df.shape[0]):
        for c in range(sheet_df.shape[1]):
            if target in _norm_txt(sheet_df.iat[r, c]):
                return r, c
    return None

def _scan_row_until_blank(sheet_df, r, start_c):
    end = start_c
    while (end < sheet_df.shape[1]
           and not pd.isna(sheet_df.iat[r, end])
           and str(sheet_df.iat[r, end]).strip() != ""):
        end += 1
    return end

def _normalize_year_cols(cols):
    m = {}
    for c in cols:
        try:
            m[int(float(str(c).strip()))] = c
        except Exception:
            pass
    return m


# ── Validation Functions ─────────────────────────────────────────────────────

def validate_main_workbook_structure(main_xls: pd.ExcelFile) -> dict:
    """
    Validates structural requirements of the main workbook:
    - Required sheets
    - Required columns in standard sheets
    - Anchors and block layouts in Sales volume parameters
    """
    res = {"ok": True, "errors": [], "warnings": []}
    
    # 1. Sheets
    conv_sheet = resolve_sheet_name(main_xls, ["conversion"])
    yield_sheet = resolve_sheet_name(main_xls, ["yield"])
    params_sheet = resolve_sheet_name(main_xls, ["product param"])
    sales_sheet = resolve_sheet_name(main_xls, ["sales", "vol"])
    
    missing_sheets = []
    if not conv_sheet: missing_sheets.append("Conversion rates (keyword 'conv')")
    if not yield_sheet: missing_sheets.append("Production yields (keyword 'yield')")
    if not params_sheet: missing_sheets.append("Product parameters (keyword 'param')")
    if not sales_sheet: missing_sheets.append("Sales volume parameters (keyword 'sales')")
    
    if missing_sheets:
        res["errors"].append(f"Missing required sheet(s): {', '.join(missing_sheets)}.")
        res["ok"] = False
        return res  # Stop early if missing sheets
        
    # 2. Columns
    # Product parameters
    params_df = main_xls.parse(params_sheet)
    params_df.columns = params_df.columns.astype(str).str.strip()
    for col in ["Parent0", "Archetype"]:
        if col not in params_df.columns:
            res["errors"].append(f"Missing column '{col}' in '{params_sheet}' sheet.")
            res["ok"] = False

    # Production yields
    yield_df = main_xls.parse(yield_sheet)
    yield_df.columns = yield_df.columns.astype(str).str.strip()
    for col in ["Parent0", "Planned yield (bu/ac)", "Actual yield"]:
        if col not in yield_df.columns:
            res["errors"].append(f"Missing column '{col}' in '{yield_sheet}' sheet.")
            res["ok"] = False

    # Conversion rates
    conv_df = main_xls.parse(conv_sheet)
    conv_df.columns = conv_df.columns.astype(str).str.strip()
    for col in ["Parent0", "totalConversionRate"]:
        if col not in conv_df.columns:
            res["errors"].append(f"Missing column '{col}' in '{conv_sheet}' sheet.")
            res["ok"] = False

    # 3. Sales block structure
    sales_raw = main_xls.parse(sales_sheet, header=None)
    
    # Median first-year block
    anchor = _find_cell(sales_raw, "Median first year sales volumes")
    if not anchor:
        res["errors"].append("Could not find the 'Median first year sales volumes' block in 'Sales volume parameters'.")
        res["ok"] = False
    else:
        anchor_r, _ = anchor
        median_header_r = None
        for r in range(anchor_r, min(anchor_r + 30, sales_raw.shape[0])):
            if _norm_txt(sales_raw.iat[r, 0]) == "archetype":
                median_header_r = r
                break
        
        if median_header_r is None:
            res["errors"].append("Couldn't find 'Archetype' header for 'Median first year sales volumes' block.")
            res["ok"] = False
        else:
            median_end_c = _scan_row_until_blank(sales_raw, median_header_r, 0)
            headers = sales_raw.iloc[median_header_r, 0:median_end_c].values
            maturity_cols_map = {}
            for col in headers:
                if _norm_txt(col) == "archetype": continue
                try:
                    maturity_cols_map[int(float(str(col).strip()))] = col
                except Exception:
                    pass
            missing_mats = [m for m in NEEDED_MATURITIES if m not in maturity_cols_map]
            if missing_mats:
                res["errors"].append(f"Missing maturity columns {missing_mats} in the 'Median first year sales volumes' block.")
                res["ok"] = False

    # Median growth block
    growth_anchor = _find_cell(sales_raw, "Median growth rates")
    if not growth_anchor:
        res["errors"].append("Could not find the 'Median growth rates' block in 'Sales volume parameters'.")
        res["ok"] = False
    else:
        ga_r, _ = growth_anchor
        growth_header_r = None
        growth_start_c = None
        for r in range(ga_r, min(ga_r + 30, sales_raw.shape[0])):
            for c in range(sales_raw.shape[1] - 1):
                if _norm_txt(sales_raw.iat[r, c]) == "archetype" and _norm_txt(sales_raw.iat[r, c + 1]) == "maturity":
                    growth_header_r = r
                    growth_start_c = c
                    break
            if growth_header_r is not None: break
            
        if growth_header_r is None:
            res["errors"].append("Couldn't find Archetype+Maturity header in 'Median growth rates' block.")
            res["ok"] = False
        else:
            growth_end_c = _scan_row_until_blank(sales_raw, growth_header_r, growth_start_c)
            headers = sales_raw.iloc[growth_header_r, growth_start_c:growth_end_c].values
            year_map = _normalize_year_cols([c for c in headers if _norm_txt(c) not in ["archetype", "maturity"]])
            missing_years = [y for y in YEARS_NEEDED if y not in year_map]
            if missing_years:
                res["errors"].append(f"Missing growth year columns {missing_years} in the 'Median growth rates' block.")
                res["ok"] = False

    return res

def extract_sales_variability(sales_xls: pd.ExcelFile) -> tuple:
    """
    Returns (sv_fy_mu, sv_fy_sig2, sv_gr_mu, sv_gr_sig2) extracted dynamically.
    Tolerates radically shifted blocks and completely missing descriptive headings
    by strictly looking for the structural geometry of the core numerical tables.
    """
    sales_sheet = resolve_sheet_name(sales_xls, ["sales", "vol"])
    if not sales_sheet: return {}, {}, {}, {}
    sv_raw = sales_xls.parse(sales_sheet, header=None)
    
    fy_rows = []
    gr_rows = []
    
    for r in range(sv_raw.shape[0]):
        for c in range(sv_raw.shape[1]):
            val = str(sv_raw.iat[r,c]).strip()
            if not val or val.lower() == 'nan': continue
            
            val_low = val.lower()
            if any(k in val_low for k in ["bayer", "syngenta", "conventional"]):
                row_nums = []
                for c2 in range(c+1, min(c+15, sv_raw.shape[1])):
                    v2 = sv_raw.iat[r, c2]
                    if pd.isna(v2) or str(v2).strip() == "": continue
                    try:
                        num = float(v2)
                        row_nums.append((c2, num))
                    except ValueError:
                        break
                
                # Check for Growth Rate structure (1 maturity + 9 year floats = 10 cols)
                if len(row_nums) >= 10:
                    mat_val = int(row_nums[0][1])
                    vals = {idx+2: v for idx, (_, v) in enumerate(row_nums[1:10])}
                    gr_rows.append((r, c, val, mat_val, vals))
                # Check for First Year structure (4 specific maturity map floats)
                elif len(row_nums) == 4:
                    vals = {85: row_nums[0][1], 95: row_nums[1][1], 105: row_nums[2][1], 115: row_nums[3][1]}
                    fy_rows.append((r, c, val, vals))
                break
                
    def group_blocks(rows_list):
        blocks = []
        if not rows_list: return blocks
        rows_list.sort(key=lambda x: x[0])
        current_block = [rows_list[0]]
        for row in rows_list[1:]:
            if row[0] - current_block[-1][0] <= 5:
                current_block.append(row)
            else:
                blocks.append(current_block)
                current_block = [row]
        blocks.append(current_block)
        return blocks

    gr_blocks = group_blocks(gr_rows)
    fy_blocks = group_blocks(fy_rows)
    
    sv_fy_mu, sv_fy_sig2, sv_gr_mu, sv_gr_sig2 = {}, {}, {}, {}
    
    if len(fy_blocks) >= 2:
        fy_blocks.sort(key=lambda b: b[0][0])
        for row in fy_blocks[0]:
            arch = row[2]
            for m, v in row[3].items(): sv_fy_mu[(arch, m)] = v
        for row in fy_blocks[1]:
            arch = row[2]
            for m, v in row[3].items(): sv_fy_sig2[(arch, m)] = v
            
    if len(gr_blocks) >= 2:
        gr_blocks.sort(key=lambda b: b[0][0])
        for row in gr_blocks[0]:
            arch, mat = row[2], row[3]
            for y, v in row[4].items(): sv_gr_mu[(arch, mat, y)] = v
        for row in gr_blocks[1]:
            arch, mat = row[2], row[3]
            for y, v in row[4].items(): sv_gr_sig2[(arch, mat, y)] = v
            
    return sv_fy_mu, sv_fy_sig2, sv_gr_mu, sv_gr_sig2

def validate_sales_variability_workbook_structure(sales_xls: pd.ExcelFile) -> dict:
    """
    Validates structural requirements of the sales variability workbook.
    """
    res = {"ok": True, "errors": [], "warnings": []}
    
    sales_sheet = resolve_sheet_name(sales_xls, ["sales", "vol"])
    if not sales_sheet:
        res["errors"].append("Missing required sheet using keyword 'sales' or 'vol'.")
        res["ok"] = False
        return res
        
    sv_fy_mu, sv_fy_sig2, sv_gr_mu, sv_gr_sig2 = extract_sales_variability(sales_xls)
    
    if not sv_fy_mu:
        res["errors"].append("Could not extract First Year Sales (mu) data block.")
        res["ok"] = False
    if not sv_fy_sig2:
        res["errors"].append("Could not extract First Year Sales (sigma^2) data block.")
        res["ok"] = False
    if not sv_gr_mu:
        res["errors"].append("Could not extract Growth Rate (mu) data block.")
        res["ok"] = False
    if not sv_gr_sig2:
        res["errors"].append("Could not extract Growth Rate (sigma^2) data block.")
        res["ok"] = False
        
    return res

def format_validation_errors(main_res: dict, sales_res: dict = None) -> list:
    """Combine results into a clean list of error strings."""
    msgs = []
    if main_res and not main_res["ok"]:
        msgs.append("Main workbook validation failed:")
        for e in main_res["errors"]:
            msgs.append(f"- {e}")
    if sales_res and not sales_res["ok"]:
        msgs.append("Sales variability workbook validation failed:")
        for e in sales_res["errors"]:
            msgs.append(f"- {e}")
    return msgs
