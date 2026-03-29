"""
Monte Carlo Production & Inventory Planner (Multi-Product)
==========================================================
Streamlit dashboard for 10-year Monte Carlo simulation of seed production
and inventory, using archetype/maturity inputs, yield & conversion
uncertainty, and portfolio-level aggregation across multiple products.

Run with:  streamlit run app.py
"""

import base64
import io
import os
import pandas as pd
import numpy as np
import streamlit as st
import altair as alt

# ── Configuration ────────────────────────────────────────────────────────────

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
MODEL_NOTEBOOK = "GDM_final (2).ipynb"

# Excel inputs — two modes:
#   local    — data.xlsx and salesVariability.xlsx next to app.py (laptop dev only).
#   secrets  — GDM_DATA_XLSX + GDM_SALES_VAR_XLSX (env or st.secrets). Streamlit Cloud
#              uses secrets mode automatically so workbooks never need to be on GitHub.
#   Override: GDM_DATA_SOURCE=local forces disk files; =secrets forces secret keys only.


def _is_streamlit_cloud() -> bool:
    """True on Streamlit Community Cloud (repo under /mount/src/); confidential Excel stays in Secrets."""
    if os.environ.get("STREAMLIT_SHARING", "").lower() == "true":
        return True
    bd = os.path.abspath(BASE_DIR).replace("\\", "/")
    return "/mount/src/" in bd


def _resolve_secret_path(key: str):
    v = os.environ.get(key)
    if v is not None and str(v).strip():
        return str(v).strip()
    try:
        if key in st.secrets:
            val = st.secrets[key]
            if val is not None and str(val).strip():
                return str(val).strip()
    except Exception:
        pass
    return None


def _config_mode() -> str:
    v = os.environ.get("GDM_DATA_SOURCE", "").strip().lower()
    if v in ("local", "secrets"):
        return v
    try:
        if "GDM_DATA_SOURCE" in st.secrets:
            s = str(st.secrets["GDM_DATA_SOURCE"]).strip().lower()
            if s in ("local", "secrets"):
                return s
    except Exception:
        pass
    if _resolve_secret_path("GDM_DATA_XLSX") and _resolve_secret_path("GDM_SALES_VAR_XLSX"):
        return "secrets"
    if _is_streamlit_cloud():
        return "secrets"
    return "local"


def _normalize_excel_input(raw: str):
    """Return a path/URL string, or BytesIO for base64-encoded .xlsx content."""
    raw = raw.strip()
    if raw.startswith("http://") or raw.startswith("https://"):
        return raw
    if os.path.isfile(raw):
        return raw
    try:
        cleaned = "".join(raw.split())
        pad = (-len(cleaned)) % 4
        if pad:
            cleaned += "=" * pad
        decoded = base64.b64decode(cleaned, validate=False)
        if len(decoded) >= 4 and decoded[:2] == b"PK":
            return io.BytesIO(decoded)
    except Exception:
        pass
    return raw


def _data_xlsx_source():
    if _config_mode() == "local":
        return os.path.join(BASE_DIR, "data.xlsx")
    raw = _resolve_secret_path("GDM_DATA_XLSX")
    if not raw:
        return None
    return _normalize_excel_input(raw)


def _sales_var_xlsx_source():
    if _config_mode() == "local":
        return os.path.join(BASE_DIR, "salesVariability.xlsx")
    raw = _resolve_secret_path("GDM_SALES_VAR_XLSX")
    if not raw:
        return None
    return _normalize_excel_input(raw)


def _describe_excel_source(src) -> str:
    if src is None:
        return "<not set>"
    if isinstance(src, io.BytesIO):
        return "<embedded workbook from secret>"
    s = str(src)
    return s if len(s) <= 200 else s[:197] + "..."


def _path_readable_for_excel(source) -> bool:
    if source is None:
        return False
    if isinstance(source, io.BytesIO):
        return True
    p = str(source)
    if p.startswith("http://") or p.startswith("https://"):
        return True
    return os.path.isfile(p)

NEEDED_MATURITIES = [85, 95, 105, 115]
YEARS_NEEDED = list(range(2, 11))
LAUNCH_YEARS = list(range(0, 6))

STRATEGY_OPTIONS = {
    "Custom multiplier by year (Y1-Y10)": "custom",
    "Just-in-time 1.0x": "jit",
    "Conservative 1.2x": "cons",
    "Aggressive 2.0x": "aggr",
}

# ── Sales-Sheet Helper Functions ─────────────────────────────────────────────

def norm_txt(x):
    if pd.isna(x):
        return ""
    return str(x).strip().lower()


def find_cell(sheet, text):
    target = text.strip().lower()
    for r in range(sheet.shape[0]):
        for c in range(sheet.shape[1]):
            if target in norm_txt(sheet.iat[r, c]):
                return r, c
    return None


def scan_row_until_blank(sheet, r, start_c):
    end = start_c
    while (end < sheet.shape[1]
           and not pd.isna(sheet.iat[r, end])
           and str(sheet.iat[r, end]).strip() != ""):
        end += 1
    return end


def clean_series(s):
    s = s.astype(str).str.strip()
    s = s.replace({"nan": np.nan, "None": np.nan, "": np.nan})
    return s


def is_bad_archetype(x):
    t = norm_txt(x)
    if t in ["", "archetype", "maturity"]:
        return True
    bad = [
        "median first year sales volumes",
        "average first year sales volumes",
        "median growth rates",
        "average growth rates",
        "relative sales year",
    ]
    return any(b in t for b in bad)


def normalize_year_cols(cols):
    m = {}
    for c in cols:
        try:
            m[int(float(str(c).strip()))] = c
        except Exception:
            pass
    return m


def to_rate(x):
    if pd.isna(x):
        return 0.0
    if isinstance(x, str):
        s = x.strip()
        if s.endswith("%"):
            s = s[:-1].strip()
            v = pd.to_numeric(s, errors="coerce")
            return 0.0 if pd.isna(v) else float(v) / 100.0
        v = pd.to_numeric(s, errors="coerce")
        if pd.isna(v):
            return 0.0
        v = float(v)
    else:
        v = float(x)
    return v / 100.0 if abs(v) > 2 else v


# ── Sales-Variability Parsing Helpers ────────────────────────────────────────

def _sv_find(sv_raw, text):
    t = text.strip().lower()
    for r in range(sv_raw.shape[0]):
        for c in range(sv_raw.shape[1]):
            v = sv_raw.iat[r, c]
            if not pd.isna(v) and t in str(v).strip().lower():
                return r, c
    return None


def _parse_sv_wide(sv_raw, header_row, start_col, mat_cols):
    result = {}
    r = header_row + 1
    while r < sv_raw.shape[0]:
        arch = sv_raw.iat[r, start_col]
        if pd.isna(arch) or str(arch).strip() == "":
            break
        arch = str(arch).strip()
        for mat, col in mat_cols.items():
            val = sv_raw.iat[r, col]
            if not pd.isna(val):
                result[(arch, int(mat))] = float(val)
        r += 1
    return result


def _parse_sv_long(sv_raw, header_row, arch_col, year_cols):
    result = {}
    r = header_row + 1
    while r < sv_raw.shape[0]:
        arch = sv_raw.iat[r, arch_col]
        mat = sv_raw.iat[r, arch_col + 1]
        if pd.isna(arch) or str(arch).strip() == "":
            break
        arch = str(arch).strip()
        if pd.isna(mat):
            r += 1
            continue
        mat = int(float(mat))
        for yr, col in year_cols.items():
            val = sv_raw.iat[r, col]
            if not pd.isna(val):
                result[(arch, mat, int(yr))] = float(val)
        r += 1
    return result


# ── Data Loading (cached) ───────────────────────────────────────────────────

@st.cache_data(show_spinner="Loading data from Excel files...")
def load_all_data():
    """Load and parse both Excel files. Returns a dict of all derived tables."""

    main_src = _data_xlsx_source()
    sales_src = _sales_var_xlsx_source()

    # -- Load sheets from data workbook (ExcelFile keeps one open file / buffer for all sheets) --
    with pd.ExcelFile(main_src) as main_xls:
        conv_tab = pd.read_excel(main_xls, sheet_name="Conversion rates")
        yield_tab = pd.read_excel(main_xls, sheet_name="Production yields")
        params_tab = pd.read_excel(main_xls, sheet_name="Product parameters")
        sales_raw = pd.read_excel(main_xls, sheet_name="Sales volume parameters", header=None)

    for df_ in [conv_tab, yield_tab, params_tab]:
        df_.columns = df_.columns.astype(str).str.strip()

    for col in ["Parent0", "Archetype"]:
        if col not in params_tab.columns:
            raise ValueError("Product parameters must contain columns: Parent0, Archetype")

    params_tab["Parent0"] = params_tab["Parent0"].astype(str).str.strip()
    params_tab["Archetype"] = params_tab["Archetype"].astype(str).str.strip()
    parent_to_arch = params_tab[["Parent0", "Archetype"]].dropna()

    # -- Yield & Conversion Prep --
    for col in ["Parent0", "Planned yield (bu/ac)", "Actual yield"]:
        if col not in yield_tab.columns:
            raise ValueError(f"Production yields sheet missing column: {col}")

    yield_tab["Parent0"] = yield_tab["Parent0"].astype(str).str.strip()
    yield_tab["Planned yield (bu/ac)"] = pd.to_numeric(yield_tab["Planned yield (bu/ac)"], errors="coerce")
    yield_tab["Actual yield"] = pd.to_numeric(yield_tab["Actual yield"], errors="coerce")
    yield_tab["Yield_Factor"] = yield_tab["Actual yield"] / yield_tab["Planned yield (bu/ac)"]
    yield_tab["Yield_Factor"] = yield_tab["Yield_Factor"].replace([np.inf, -np.inf], np.nan)
    yield_w_arch = yield_tab.merge(parent_to_arch, on="Parent0", how="left")

    for col in ["Parent0", "totalConversionRate"]:
        if col not in conv_tab.columns:
            raise ValueError(f"Conversion rates sheet missing column: {col}")
    conv_tab["Parent0"] = conv_tab["Parent0"].astype(str).str.strip()
    conv_tab["totalConversionRate"] = pd.to_numeric(conv_tab["totalConversionRate"], errors="coerce")
    conv_w_arch = conv_tab.merge(parent_to_arch, on="Parent0", how="left")

    # -- Median First-Year Sales --
    anchor = find_cell(sales_raw, "Median first year sales volumes")
    if anchor is None:
        raise ValueError("Couldn't find 'Median first year sales volumes' block.")
    anchor_r, _ = anchor

    median_header_r = None
    for r in range(anchor_r, min(anchor_r + 30, sales_raw.shape[0])):
        if norm_txt(sales_raw.iat[r, 0]) == "archetype":
            median_header_r = r
            break
    if median_header_r is None:
        raise ValueError("Couldn't find 'Archetype' header for Median first year sales block.")

    median_end_c = scan_row_until_blank(sales_raw, median_header_r, 0)
    median_df = sales_raw.iloc[median_header_r + 1:, 0:median_end_c].copy()
    median_df.columns = sales_raw.iloc[median_header_r, 0:median_end_c].values
    median_df = median_df.dropna(subset=["Archetype"])
    median_df["Archetype"] = clean_series(median_df["Archetype"])
    median_df = median_df[~median_df["Archetype"].apply(is_bad_archetype)]

    maturity_cols_map = {}
    for col in median_df.columns:
        if norm_txt(col) == "archetype":
            continue
        try:
            maturity_cols_map[int(float(str(col).strip()))] = col
        except Exception:
            pass

    missing = [m for m in NEEDED_MATURITIES if m not in maturity_cols_map]
    if missing:
        raise ValueError(
            f"Missing maturity cols {missing} in median sales block. "
            f"Found: {sorted(maturity_cols_map.keys())}"
        )

    median_sales_df = median_df[["Archetype"] + [maturity_cols_map[m] for m in NEEDED_MATURITIES]].copy()
    median_sales_df.columns = ["Archetype"] + NEEDED_MATURITIES
    for m in NEEDED_MATURITIES:
        median_sales_df[m] = pd.to_numeric(median_sales_df[m], errors="coerce")
    median_sales_df = median_sales_df.dropna(subset=NEEDED_MATURITIES, how="all")

    # -- Median Growth Rates --
    growth_anchor = find_cell(sales_raw, "Median growth rates")
    if growth_anchor is None:
        raise ValueError("Couldn't find 'Median growth rates' block.")
    ga_r, _ = growth_anchor

    growth_header_r = None
    growth_start_c = None
    for r in range(ga_r, min(ga_r + 30, sales_raw.shape[0])):
        for c in range(sales_raw.shape[1] - 1):
            if norm_txt(sales_raw.iat[r, c]) == "archetype" and norm_txt(sales_raw.iat[r, c + 1]) == "maturity":
                growth_header_r = r
                growth_start_c = c
                break
        if growth_header_r is not None:
            break
    if growth_header_r is None:
        raise ValueError("Couldn't find Archetype+Maturity header in Median growth rates block.")

    growth_end_c = scan_row_until_blank(sales_raw, growth_header_r, growth_start_c)
    growth_df = sales_raw.iloc[growth_header_r + 1:, growth_start_c:growth_end_c].copy()
    growth_df.columns = sales_raw.iloc[growth_header_r, growth_start_c:growth_end_c].values
    growth_df = growth_df.dropna(subset=["Archetype", "Maturity"])
    growth_df["Archetype"] = clean_series(growth_df["Archetype"])
    growth_df["Maturity"] = pd.to_numeric(growth_df["Maturity"], errors="coerce")
    growth_df = growth_df[~growth_df["Archetype"].apply(is_bad_archetype)]
    growth_df = growth_df.dropna(subset=["Archetype", "Maturity"])

    year_map = normalize_year_cols(
        [c for c in growth_df.columns if norm_txt(c) not in ["archetype", "maturity"]]
    )
    missing_years = [y for y in YEARS_NEEDED if y not in year_map]
    if missing_years:
        raise ValueError(
            f"Missing growth year columns {missing_years} in Median growth rates block. "
            f"Found: {sorted(year_map.keys())}"
        )

    # -- Sales Variability (lognormal params) --
    # TODO: Row/column indices here are hardcoded to the current salesVariability.xlsx layout.
    #       If that spreadsheet is restructured, these offsets will need updating.
    sv_raw = pd.read_excel(sales_src, sheet_name="Sales volume parameters", header=None)

    fy_mat_cols = {85: 1, 95: 2, 105: 3, 115: 4}
    fy_mu_header = 4
    sv_fy_mu = _parse_sv_wide(sv_raw, fy_mu_header, 0, fy_mat_cols)

    sig2_anchor = _sv_find(sv_raw, "Log-normal (sigma^2)")
    if sig2_anchor is None:
        raise ValueError("Could not find 'Log-normal (sigma^2)' block in salesVariability.xlsx")
    sig2_header = sig2_anchor[0] + 2
    sv_fy_sig2 = _parse_sv_wide(sv_raw, sig2_header, 0, fy_mat_cols)

    gr_mu_arch_col = 7
    gr_mu_year_cols = {yr: 9 + (yr - 2) for yr in range(2, 11)}
    sv_gr_mu = _parse_sv_long(sv_raw, 4, gr_mu_arch_col, gr_mu_year_cols)

    gr_sig2_arch_col = 24
    gr_sig2_year_cols = {yr: 26 + (yr - 2) for yr in range(2, 11)}
    sv_gr_sig2 = _parse_sv_long(sv_raw, 4, gr_sig2_arch_col, gr_sig2_year_cols)

    # -- Build product options list --
    product_options = []
    for a in sorted(median_sales_df["Archetype"].dropna().unique()):
        for m in NEEDED_MATURITIES:
            row = median_sales_df[median_sales_df["Archetype"] == a]
            if not row.empty:
                v = row[m].dropna()
                if not v.empty and not np.isnan(float(v.iloc[0])):
                    product_options.append(f"{a} | {m}")

    return {
        "median_sales_df": median_sales_df,
        "growth_df": growth_df,
        "year_map": year_map,
        "yield_w_arch": yield_w_arch,
        "conv_w_arch": conv_w_arch,
        "sv_fy_mu": sv_fy_mu,
        "sv_fy_sig2": sv_fy_sig2,
        "sv_gr_mu": sv_gr_mu,
        "sv_gr_sig2": sv_gr_sig2,
        "product_options": product_options,
    }


# ── Lookup & Sampling Functions ──────────────────────────────────────────────

def get_median_sales(data, archetype, maturity):
    """Fallback: return median first-year sales from data.xlsx."""
    row = data["median_sales_df"][data["median_sales_df"]["Archetype"] == archetype]
    if row.empty:
        return None
    v = row[maturity].dropna()
    return None if v.empty else float(v.iloc[0])


def get_yoy_rates(data, archetype, maturity):
    """Fallback: return median YoY growth rates from data.xlsx."""
    gdf = data["growth_df"]
    ymap = data["year_map"]
    row = gdf[(gdf["Archetype"] == archetype) & (gdf["Maturity"] == maturity)]
    if row.empty:
        return None
    return [to_rate(row[ymap[y]].iloc[0]) for y in YEARS_NEEDED]


def get_mean_std_by_archetype(data, archetype):
    """Return mean and std of yield factor and conversion rate for archetype."""
    y = data["yield_w_arch"].loc[data["yield_w_arch"]["Archetype"] == archetype, "Yield_Factor"].dropna()
    c = data["conv_w_arch"].loc[data["conv_w_arch"]["Archetype"] == archetype, "totalConversionRate"].dropna()
    y_mean = float(y.mean()) if len(y) else float(data["yield_w_arch"]["Yield_Factor"].dropna().mean())
    c_mean = float(c.mean()) if len(c) else float(data["conv_w_arch"]["totalConversionRate"].dropna().mean())
    y_std = float(y.std(ddof=1)) if len(y) > 1 else 1e-6
    c_std = float(c.std(ddof=1)) if len(c) > 1 else 1e-6
    if y_std <= 0:
        y_std = 1e-6
    if c_std <= 0:
        c_std = 1e-6
    return y_mean, y_std, c_mean, c_std


def sample_yield_conv_normal(data, archetype, rng):
    """Draw one yield factor and one conversion rate for this run."""
    y_mean, y_std, c_mean, c_std = get_mean_std_by_archetype(data, archetype)
    return (
        float(max(0.0, rng.normal(y_mean, y_std))),
        float(max(0.0, rng.normal(c_mean, c_std))),
    )


def build_sales_curve(data, archetype, maturity):
    """
    Returns a fixed 10-year sales array using median values from data.xlsx.
    Used only as a fallback when lognormal parameters are unavailable.
    """
    y1 = get_median_sales(data, archetype, maturity)
    yoy = get_yoy_rates(data, archetype, maturity)
    if y1 is None or yoy is None:
        return None
    sales = [y1]
    for rate in yoy:
        sales.append(max(0.0, sales[-1] * (1 + rate)))
    return np.array(sales, dtype=float)


def draw_sales_curve(data, archetype, maturity, rng):
    """
    Draw one complete 10-year sales trajectory from lognormal parameters.

    Year 1 is sampled from Lognormal(mu, sigma). Each subsequent year's value
    is computed by multiplying the prior year by a growth multiplier drawn from
    Lognormal(mu, sigma) for that year.

    Growth rate parameters near -9.21 (= log(~0)) signal end-of-lifecycle;
    the drawn multiplier will be near zero, naturally collapsing sales.

    Falls back to build_sales_curve() if any parameter is missing.
    """
    fy_mu = data["sv_fy_mu"].get((archetype, maturity))
    fy_sig2 = data["sv_fy_sig2"].get((archetype, maturity))

    if fy_mu is None or fy_sig2 is None:
        return build_sales_curve(data, archetype, maturity)

    fy_sigma = float(np.sqrt(max(fy_sig2, 1e-10)))
    y1 = float(rng.lognormal(fy_mu, fy_sigma))

    sales = [y1]
    for yr in range(2, 11):
        gr_mu = data["sv_gr_mu"].get((archetype, maturity, yr))
        gr_sig2 = data["sv_gr_sig2"].get((archetype, maturity, yr))

        if gr_mu is None or gr_sig2 is None:
            yoy = get_yoy_rates(data, archetype, maturity)
            if yoy is not None:
                rate = yoy[yr - 2]
                sales.append(max(0.0, sales[-1] * (1.0 + rate)))
            else:
                sales.append(0.0)
            continue

        gr_sigma = float(np.sqrt(max(gr_sig2, 1e-10)))
        growth_draw = float(rng.lognormal(gr_mu, gr_sigma))
        next_sales = sales[-1] * growth_draw
        sales.append(max(0.0, next_sales))

    return np.array(sales, dtype=float)


# ── Simulation Engine ────────────────────────────────────────────────────────

def get_yearly_multipliers(strategy, custom_sliders):
    """Return list of 10 production multipliers (Y1..Y10)."""
    if strategy == "custom":
        return list(custom_sliders)
    elif strategy == "jit":
        return [1.0] * 10
    elif strategy == "cons":
        return [1.2] * 10
    elif strategy == "aggr":
        return [2.0] * 10
    return [1.5] * 10


def simulate_one_run(data, sales, archetype, rng, mults, use_floor, min_floor,
                     use_max_carry, max_carryover):
    """
    One Monte Carlo lifecycle:
      sales      -- length-10 array (Year 1..10)
      mults      -- 10 production multipliers
      use_floor  -- whether to enforce minimum production floor
      min_floor  -- minimum production floor value
      use_max_carry -- whether to enforce maximum carryover cap
      max_carryover -- absolute carry-in inventory cap
    """
    y_draw, c_draw = sample_yield_conv_normal(data, archetype, rng)

    carryover = 0.0
    rows = []
    missed_sales = []

    for yr in range(10):
        expected_sales_next = sales[yr + 1] if yr < 9 else 0.0

        planned_prod = mults[yr] * expected_sales_next

        if use_max_carry and max_carryover > 0.0 and carryover >= max_carryover:
            planned_prod = 0.0

        if use_floor and planned_prod > 0.0:
            planned_prod = max(planned_prod, min_floor)

        new_prod = planned_prod * y_draw * c_draw

        prod_loss = new_prod * 0.02
        carry_loss = carryover * 0.10

        total_saleable = (carryover - carry_loss) + (new_prod - prod_loss)
        remaining = total_saleable - sales[yr]

        missed = max(0.0, -remaining)
        missed_sales.append(missed)

        rows.append([
            carryover,
            -carry_loss,
            planned_prod,
            new_prod,
            -prod_loss,
            total_saleable,
            sales[yr],
            remaining,
        ])
        carryover = remaining

    return np.array(rows, dtype=float), np.array(missed_sales, dtype=float)


def determine_analysis_year(lifecycle_df, year_mode, custom_year_idx):
    """Resolve year-mode selection to a 0-based year index."""
    sales_row = lifecycle_df.loc["Sales"].astype(float).values
    if year_mode == "last_sales":
        idx = np.where(sales_row > 0)[0]
        return int(idx[-1]) if len(idx) else 9
    return int(custom_year_idx)


def build_lifecycle_sim(data, products, iterations, seed, strategy, custom_sliders,
                        use_floor, min_floor, use_max_carry, max_carryover,
                        year_mode, custom_year_idx, threshold):
    """
    Full Monte Carlo runner.

    Returns (lifecycle_df, summary_df, ay_idx, parsed_products) or None on error.
    """
    parsed = []
    warnings = []
    for p in products:
        arch, mat_str = p.split("|")
        arch = arch.strip()
        maturity = int(mat_str.strip())
        if build_sales_curve(data, arch, maturity) is None:
            warnings.append(f"Missing sales data for {arch} | {maturity}; skipping.")
            continue
        parsed.append((arch, maturity))

    if not parsed:
        return None, None, None, None, ["No valid products selected."]

    mults = get_yearly_multipliers(strategy, custom_sliders)
    rng = np.random.default_rng(int(seed))
    run_rows_all = []

    for arch, maturity in parsed:
        prod_rows = []
        for _ in range(int(iterations)):
            sales = draw_sales_curve(data, arch, maturity, rng)
            rows, _ = simulate_one_run(
                data, sales, arch, rng, mults,
                use_floor, min_floor, use_max_carry, max_carryover
            )
            prod_rows.append(rows)
        prod_rows = np.stack(prod_rows, axis=0)
        run_rows_all.append(prod_rows)

    run_rows = np.sum(run_rows_all, axis=0)
    mean_rows = run_rows.mean(axis=0)

    cols = [f"Year {i}" for i in range(1, 11)]
    lifecycle_df = pd.DataFrame(
        mean_rows.T,
        columns=cols,
        index=[
            "Carry-in inventory (from prior year)",
            "Carry-in quality loss (10%)",
            "Planned production",
            "Actual production (after yield & conversion)",
            "Production quality loss (2%)",
            "Total saleable inventory",
            "Sales",
            "Remaining inventory [negative = stockout]",
        ],
    )

    remaining_all = run_rows[:, :, 7]
    ay_idx = determine_analysis_year(lifecycle_df, year_mode, custom_year_idx)
    thr = float(threshold)

    sel_rem = remaining_all[:, ay_idx]
    end_rem = remaining_all[:, 9]

    def _row_stats(arr):
        return {
            "Mean remaining": float(arr.mean()),
            "Median remaining": float(np.median(arr)),
            "P90 remaining": float(np.percentile(arr, 90)),
            "P(remaining > 0)": float((arr > 0).mean()),
            f"P(remaining > {thr:.0f})": float((arr > thr).mean()),
            "P(stockout)": float((arr < 0).mean()),
        }

    selected_stats = _row_stats(sel_rem)
    end_stats = _row_stats(end_rem)

    summary_df = pd.DataFrame.from_dict(
        {
            f"Selected year (Year {ay_idx + 1})": selected_stats,
            "End of lifecycle (Year 10)": end_stats,
        },
        orient="index",
    )

    return lifecycle_df, summary_df, ay_idx, parsed, warnings


def build_launch_cohorts(launch_plan_df):
    """From launch plan table, create (archetype, maturity, launch_year, n_products)."""
    cohorts = []
    for _, row in launch_plan_df.iterrows():
        arch = str(row["Archetype"]).strip()
        maturity = int(row["Maturity"])
        for y in LAUNCH_YEARS:
            n = int(row.get(f"Y{y}", 0) or 0)
            if n > 0:
                cohorts.append((arch, maturity, y, n))
    return cohorts


def run_multiyear_launch_sim(data, launch_plan_df, iterations, seed, strategy, custom_sliders,
                             use_floor, min_floor, use_max_carry, max_carryover, threshold):
    """Multi-year launch Monte Carlo aggregated over calendar years."""
    cohorts = build_launch_cohorts(launch_plan_df)
    if not cohorts:
        return None, None, None, ["No launches defined. Fill in at least one launch cell > 0."]

    max_launch = max(c[2] for c in cohorts)
    horizon_years = max_launch + 10
    rng = np.random.default_rng(int(seed))
    mults = get_yearly_multipliers(strategy, custom_sliders)

    all_runs = np.zeros((int(iterations), horizon_years, 8), dtype=float)
    unique_products = sorted({(arch, mat) for (arch, mat, _, _) in cohorts})

    for arch, mat in unique_products:
        base_runs = []
        for _ in range(int(iterations)):
            sales = draw_sales_curve(data, arch, mat, rng)
            rows, _ = simulate_one_run(
                data, sales, arch, rng, mults,
                use_floor, min_floor, use_max_carry, max_carryover
            )
            base_runs.append(rows)
        base_runs = np.stack(base_runs, axis=0)

        prod_all = np.zeros((int(iterations), horizon_years, 8), dtype=float)
        for _, _, launch_year, n_products in [c for c in cohorts if c[0] == arch and c[1] == mat]:
            for _ in range(n_products):
                for it in range(int(iterations)):
                    rows = base_runs[it, :, :]
                    start = launch_year
                    end = min(launch_year + 10, horizon_years)
                    span = end - start
                    prod_all[it, start:end, :] += rows[:span, :]
        all_runs += prod_all

    median_rows = np.median(all_runs, axis=0)
    year_cols = [f"Year {i}" for i in range(1, horizon_years + 1)]
    lifecycle_df = pd.DataFrame(
        median_rows.T,
        columns=year_cols,
        index=[
            "Carry-in inventory (from prior year)",
            "Carry-in quality loss (10%)",
            "Planned production",
            "Actual production (after yield & conversion)",
            "Production quality loss (2%)",
            "Total saleable inventory",
            "Sales",
            "Remaining inventory [negative = stockout]",
        ],
    )

    remaining_all = all_runs[:, :, 7]
    sales_row = lifecycle_df.loc["Sales"].astype(float).values
    s_idx = np.where(sales_row > 0)[0]
    ay_idx = int(s_idx[-1]) if len(s_idx) else horizon_years - 1

    thr = float(threshold)

    def _row_stats(arr):
        return {
            "Mean remaining": float(arr.mean()),
            "Median remaining": float(np.median(arr)),
            "P90 remaining": float(np.percentile(arr, 90)),
            "P(remaining > 0)": float((arr > 0).mean()),
            f"P(remaining > {thr:.0f})": float((arr > thr).mean()),
            "P(stockout)": float((arr < 0).mean()),
        }

    summary_df = pd.DataFrame.from_dict(
        {
            f"Last sales year (Year {ay_idx + 1})": _row_stats(remaining_all[:, ay_idx]),
            f"End of horizon (Year {horizon_years})": _row_stats(remaining_all[:, -1]),
        },
        orient="index",
    )
    return lifecycle_df, summary_df, ay_idx, []


# ── Streamlit Dashboard ──────────────────────────────────────────────────────

def render_sidebar(data):
    """Render all sidebar controls and return their current values as a dict."""

    st.sidebar.header("Simulation Mode")
    mode_label = st.sidebar.radio(
        "Mode",
        options=["Single-start portfolio", "Multi-year launch cohorts"],
        index=0,
    )
    mode = "single" if mode_label == "Single-start portfolio" else "multi"

    st.sidebar.header("Product Selection")
    selected_products = st.sidebar.multiselect(
        "Products (Archetype | Maturity)",
        options=data["product_options"],
        default=[data["product_options"][0]] if data["product_options"] else [],
    )

    st.sidebar.header("Production Strategy")
    strategy_label = st.sidebar.selectbox(
        "Strategy",
        options=list(STRATEGY_OPTIONS.keys()),
        index=0,
    )
    strategy = STRATEGY_OPTIONS[strategy_label]

    custom_sliders = [1.5] * 10
    if strategy == "custom":
        st.sidebar.markdown("**Per-year production multipliers**")
        cols_left, cols_right = st.sidebar.columns(2)
        for i in range(10):
            target = cols_left if i < 5 else cols_right
            custom_sliders[i] = target.slider(
                f"Y{i + 1}", min_value=0.5, max_value=3.0, value=1.5, step=0.1,
                key=f"mult_y{i + 1}",
            )

    st.sidebar.header("Simulation Settings")
    iterations = st.sidebar.slider("Number of simulations", 100, 5000, 1000, step=100)
    seed = st.sidebar.number_input("Random seed", value=42, step=1)

    st.sidebar.header("Analysis Year")
    year_mode_label = st.sidebar.selectbox(
        "Year mode",
        options=["Last Year of Sales", "Choose specific year"],
    )
    year_mode = "last_sales" if year_mode_label == "Last Year of Sales" else "custom"
    custom_year_idx = 4
    if year_mode == "custom":
        custom_year = st.sidebar.selectbox(
            "Analysis year",
            options=[f"Year {i}" for i in range(1, 11)],
            index=4,
        )
        custom_year_idx = int(custom_year.split()[-1]) - 1

    threshold = st.sidebar.slider(
        "Remaining inventory threshold",
        min_value=-10000.0, max_value=50000.0, value=0.0, step=500.0,
    )

    st.sidebar.header("Production Constraints")
    use_floor = st.sidebar.checkbox("Enable minimum production floor", value=False)
    min_floor = 0.0
    if use_floor:
        min_floor = st.sidebar.slider(
            "Min production floor",
            min_value=0.0, max_value=20000.0, value=0.0, step=500.0,
        )
    use_max_carry = st.sidebar.checkbox("Enable maximum carryover cap", value=False)
    max_carryover = 0.0
    if use_max_carry:
        max_carryover = st.sidebar.slider(
            "Max carry-in inventory",
            min_value=0.0, max_value=50000.0, value=10000.0, step=500.0,
        )

    launch_plan_df = None
    if mode == "multi":
        st.sidebar.header("Launch Cohorts (Y0-Y5)")
        launch_rows = []
        for a in sorted(data["median_sales_df"]["Archetype"].dropna().unique()):
            for m in NEEDED_MATURITIES:
                launch_rows.append({
                    "Archetype": a,
                    "Maturity": m,
                    **{f"Y{y}": 0 for y in LAUNCH_YEARS},
                })
        launch_plan_df = st.sidebar.data_editor(
            pd.DataFrame(launch_rows),
            num_rows="fixed",
            use_container_width=True,
            key="launch_plan_editor",
        )

    return {
        "mode": mode,
        "mode_label": mode_label,
        "products": selected_products,
        "strategy": strategy,
        "strategy_label": strategy_label,
        "custom_sliders": custom_sliders,
        "iterations": iterations,
        "seed": seed,
        "year_mode": year_mode,
        "year_mode_label": year_mode_label,
        "custom_year_idx": custom_year_idx,
        "threshold": threshold,
        "use_floor": use_floor,
        "min_floor": min_floor,
        "use_max_carry": use_max_carry,
        "max_carryover": max_carryover,
        "launch_plan_df": launch_plan_df,
    }


def _render_chart_view(lifecycle_df, year_order):
    """Render the Chart View tab with three visualizations."""

    # ── 1. Sales vs Production by Year (grouped bar) ─────────────────────
    st.subheader("Sales vs Production by Year")
    metrics = {
        "Planned Production": lifecycle_df.loc["Planned production"].values,
        "Actual Production": lifecycle_df.loc["Actual production (after yield & conversion)"].values,
        "Sales": lifecycle_df.loc["Sales"].values,
    }
    sp_df = pd.DataFrame([
        {"Year": yr, "Metric": metric, "Value": float(vals[i])}
        for metric, vals in metrics.items()
        for i, yr in enumerate(year_order)
    ])
    sp_chart = alt.Chart(sp_df).mark_bar().encode(
        x=alt.X("Year:N", sort=year_order, title="Year"),
        y=alt.Y("Value:Q", title="Units"),
        color=alt.Color("Metric:N", scale=alt.Scale(
            domain=["Planned Production", "Actual Production", "Sales"],
            range=["#4c78a8", "#72b7b2", "#f58518"],
        )),
        xOffset="Metric:N",
    )
    st.altair_chart(sp_chart, use_container_width=True)

    # ── 2. Total Saleable Inventory vs Sales (overlay line) ──────────────
    st.subheader("Total Saleable Inventory vs Sales")
    line_metrics = {
        "Total Saleable Inventory": lifecycle_df.loc["Total saleable inventory"].values,
        "Sales": lifecycle_df.loc["Sales"].values,
    }
    line_df = pd.DataFrame([
        {"Year": yr, "Metric": metric, "Value": float(vals[i])}
        for metric, vals in line_metrics.items()
        for i, yr in enumerate(year_order)
    ])
    line_chart = alt.Chart(line_df).mark_line(point=True, strokeWidth=2.5).encode(
        x=alt.X("Year:N", sort=year_order, title="Year"),
        y=alt.Y("Value:Q", title="Units"),
        color=alt.Color("Metric:N", scale=alt.Scale(
            domain=["Total Saleable Inventory", "Sales"],
            range=["#4c78a8", "#f58518"],
        )),
    )
    st.altair_chart(line_chart, use_container_width=True)

    # ── 3. Inventory Flow Waterfall ──────────────────────────────────────
    st.subheader("Inventory Flow Waterfall")
    wf_year = st.selectbox(
        "Select year",
        options=year_order,
        index=0,
        key="waterfall_year_select",
    )
    col_data = lifecycle_df[wf_year]

    steps = [
        ("Carry-in",          float(col_data.loc["Carry-in inventory (from prior year)"])),
        ("Carry-in Loss",     float(col_data.loc["Carry-in quality loss (10%)"])),
        ("Actual Production", float(col_data.loc["Actual production (after yield & conversion)"])),
        ("Production Loss",   float(col_data.loc["Production quality loss (2%)"])),
        ("Total Saleable",    None),
        ("Sales",             float(-col_data.loc["Sales"])),
        ("Remaining",         None),
    ]

    running = 0.0
    wf_rows = []
    for label, delta in steps:
        if label == "Total Saleable":
            wf_rows.append({"Step": label, "Start": 0, "End": running, "Type": "total"})
        elif label == "Remaining":
            wf_rows.append({"Step": label, "Start": 0, "End": running, "Type": "total"})
        else:
            start = running
            running += delta
            wf_rows.append({
                "Step": label,
                "Start": start,
                "End": running,
                "Type": "increase" if delta >= 0 else "decrease",
            })

    wf_df = pd.DataFrame(wf_rows)
    step_order = [r["Step"] for r in wf_rows]

    wf_chart = alt.Chart(wf_df).mark_bar(size=40).encode(
        x=alt.X("Step:N", sort=step_order, title=None),
        y=alt.Y("Start:Q", title="Units"),
        y2="End:Q",
        color=alt.Color("Type:N", scale=alt.Scale(
            domain=["increase", "decrease", "total"],
            range=["#72b7b2", "#e45756", "#4c78a8"],
        ), legend=alt.Legend(title="Flow")),
    )
    st.altair_chart(wf_chart, use_container_width=True)


def render_results(lifecycle_df, summary_df, ay_idx, parsed, warnings, params):
    """Render simulation results in the main area."""

    # Warnings
    for w in (warnings or []):
        st.warning(w)

    if lifecycle_df is None:
        st.error("No valid products selected. Please choose at least one product.")
        return

    # Config summary (always visible above tabs)
    product_list_str = ", ".join([f"{a} | {m}" for a, m in parsed]) if parsed else "Launch cohort mix"
    st.markdown(f"""
**Products:** {product_list_str}
| Setting | Value |
|---|---|
| Mode | {params['mode_label']} |
| Source model | {MODEL_NOTEBOOK} |
| Strategy | {params['strategy_label']} |
| Distribution | Lognormal (salesVariability.xlsx) &mdash; Yield & conversion: Normal |
| Simulations | {params['iterations']:,} |
| Random seed | {params['seed']} |
| Analysis year | Year {ay_idx + 1} ({params['year_mode_label']}) |
| Threshold | {params['threshold']:,.0f} units |
""")

    # Key Metrics cards (always visible above tabs)
    st.subheader(f"Key Metrics — Year {ay_idx + 1}")
    sel_row = summary_df.iloc[0]
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Mean Remaining", f"{sel_row['Mean remaining']:,.0f}")
    c2.metric("Median Remaining", f"{sel_row['Median remaining']:,.0f}")
    c3.metric("P(remaining > 0)", f"{sel_row['P(remaining > 0)']:.1%}")
    c4.metric("P(stockout)", f"{sel_row['P(stockout)']:.1%}")

    year_order = list(lifecycle_df.columns)

    if "view_mode" not in st.session_state:
        st.session_state["view_mode"] = "table"

    btn_left, btn_right, _ = st.columns([1, 1, 4])
    with btn_left:
        if st.button("Table View", use_container_width=True,
                      type="primary" if st.session_state["view_mode"] == "table" else "secondary"):
            st.session_state["view_mode"] = "table"
            st.rerun()
    with btn_right:
        if st.button("Chart View", use_container_width=True,
                      type="primary" if st.session_state["view_mode"] == "chart" else "secondary"):
            st.session_state["view_mode"] = "chart"
            st.rerun()

    st.divider()

    if st.session_state["view_mode"] == "table":
        st.subheader("Mean Lifecycle Across All Simulations")
        st.dataframe(lifecycle_df.round(1), use_container_width=True)

        st.subheader("Inventory Probability Summary")
        st.dataframe(summary_df.round(3), use_container_width=True)

        st.subheader("Mean Remaining Inventory by Year")
        remaining_row = lifecycle_df.loc["Remaining inventory [negative = stockout]"]
        chart_df = pd.DataFrame({
            "Year": year_order,
            "Remaining Inventory": remaining_row.values,
        })
        chart = alt.Chart(chart_df).mark_bar().encode(
            x=alt.X("Year:N", sort=year_order),
            y=alt.Y("Remaining Inventory:Q"),
        )
        st.altair_chart(chart, use_container_width=True)
    else:
        _render_chart_view(lifecycle_df, year_order)


# ── Main ─────────────────────────────────────────────────────────────────────

def main():
    st.set_page_config(
        page_title="Monte Carlo Inventory Planner",
        page_icon="📦",
        layout="wide",
    )

    title_col, btn_col = st.columns([4, 1])
    with title_col:
        st.title("Monte Carlo Production & Inventory Planner")
        st.caption("10-year lifecycle simulation with lognormal sales variability, yield & conversion uncertainty")

    data_src = _data_xlsx_source()
    sv_src = _sales_var_xlsx_source()
    if not _path_readable_for_excel(data_src) or not _path_readable_for_excel(sv_src):
        if _config_mode() == "secrets":
            st.error("Workbook secrets are missing, unreadable, or the URLs could not be loaded.")
            st.markdown(
                "On **Streamlit Cloud**, open your app → **⋮** → **Settings** → **Secrets** and "
                "define **both** keys exactly (names are case-sensitive): **`GDM_DATA_XLSX`** and "
                "**`GDM_SALES_VAR_XLSX`**. Paste a direct **`https://`** link to each file, or "
                "paste **base64**-encoded `.xlsx` content. You do **not** need to commit Excel "
                "files to GitHub. See `.streamlit/secrets.toml.example`."
            )
        else:
            st.error("Required Excel files were not found.")
        if not _path_readable_for_excel(data_src):
            st.markdown(
                f"- **Main workbook** (`data.xlsx` sheets): `{_describe_excel_source(data_src)}`"
            )
        if not _path_readable_for_excel(sv_src):
            st.markdown(
                f"- **Sales variability** (`salesVariability.xlsx`): `{_describe_excel_source(sv_src)}`"
            )
        if _config_mode() != "secrets":
            st.markdown(
                "In **local** mode, place **`data.xlsx`** and **`salesVariability.xlsx`** next to "
                "`app.py`. For Streamlit Cloud, the app uses **secrets** mode there automatically."
            )
        st.stop()

    data = load_all_data()
    params = render_sidebar(data)

    with btn_col:
        st.markdown("<div style='height: 1.2rem'></div>", unsafe_allow_html=True)
        run_clicked = st.button("Run Simulation", type="primary", use_container_width=True)

    if params["mode"] == "single" and not params["products"]:
        st.info("Select one or more products from the sidebar, then click **Run Simulation**.")
        return

    if run_clicked:
        with st.spinner("Running Monte Carlo simulation..."):
            if params["mode"] == "single":
                lifecycle_df, summary_df, ay_idx, parsed, warnings = build_lifecycle_sim(
                    data,
                    params["products"],
                    params["iterations"],
                    params["seed"],
                    params["strategy"],
                    params["custom_sliders"],
                    params["use_floor"],
                    params["min_floor"],
                    params["use_max_carry"],
                    params["max_carryover"],
                    params["year_mode"],
                    params["custom_year_idx"],
                    params["threshold"],
                )
            else:
                lifecycle_df, summary_df, ay_idx, warnings = run_multiyear_launch_sim(
                    data,
                    params["launch_plan_df"],
                    params["iterations"],
                    params["seed"],
                    params["strategy"],
                    params["custom_sliders"],
                    params["use_floor"],
                    params["min_floor"],
                    params["use_max_carry"],
                    params["max_carryover"],
                    params["threshold"],
                )
                parsed = []
        st.session_state["results"] = {
            "lifecycle_df": lifecycle_df,
            "summary_df": summary_df,
            "ay_idx": ay_idx,
            "parsed": parsed,
            "warnings": warnings,
            "params": params,
        }

    if "results" in st.session_state:
        r = st.session_state["results"]
        render_results(
            r["lifecycle_df"],
            r["summary_df"],
            r["ay_idx"],
            r["parsed"],
            r["warnings"],
            r["params"],
        )


if __name__ == "__main__":
    main()
