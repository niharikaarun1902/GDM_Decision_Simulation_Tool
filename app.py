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
import json
import time
import requests
import pandas as pd
import numpy as np
import streamlit as st
import altair as alt

# ── LLM Helpers ──────────────────────────────────────────────────────────────

ENABLE_LLM_EXPLAIN = True

BASE_URL = "https://genai.rcac.purdue.edu/api/chat/completions"
MODEL = "llama3.1:latest"  # or the exact Purdue model id
API_KEY = "sk-54fe85005f2f41a69bb5b9d56c46c1d3"  # your Purdue key

def call_llm_api(payload, attempts=3):
    headers = {
        "Authorization": f"Bearer {API_KEY}",
        "Content-Type": "application/json",
    }

    for attempt in range(attempts):
        resp = requests.post(BASE_URL, headers=headers, json=payload, timeout=30)
        if resp.status_code == 200:
            data = resp.json()
            return data["choices"][0]["message"]["content"]
        if resp.status_code == 429 and attempt < attempts - 1:
            time.sleep(2 ** attempt)
            continue
        raise RuntimeError(f"API call failed with status {resp.status_code}. {resp.text}")

def explain_table(df, context_text=""):
    """Builds a JSON payload from a DataFrame and renders an LLM explanation."""
    if not ENABLE_LLM_EXPLAIN:
        return

    st.write("DEBUG: explain_table called", {"enabled": ENABLE_LLM_EXPLAIN, "context": context_text})

    try:
        # Build compact payload representation of the dataframe
        df_json = df.to_json(orient="split")
        prompt = (
            "Please provide a very concise, professional interpretation of this "
            "Monte Carlo simulation table.\n"
            f"Context: {context_text}\n"
            f"Table data: {df_json}\n"
            "IMPORTANT: All quantities are in bushels (units of seed), NOT dollars. "
            "Do not use dollar signs ($) or currency symbols anywhere in your response. "
            "Use 'units' or 'bushels' when referring to quantities. "
            "Focus on key takeaways and risks (e.g., stockouts, excess inventory)."
        )

        payload = {
            "model": MODEL,
            "messages": [
                {"role": "system", "content": "You are a helpful data analyst."},
                {"role": "user", "content": prompt},
            ],
            "max_tokens": 200,
            "temperature": 0.4,
        }

        st.write("DEBUG: about to call LLM", {"model": MODEL, "base_url": BASE_URL})
        with st.spinner("Generating AI explanation..."):
            explanation = call_llm_api(payload, attempts=1)

        st.markdown(f"**AI interpretation:**\n{explanation}")
    except Exception as e:
        st.error(f"LLM explanation error: {repr(e)}")
        st.stop()

# ── Configuration ────────────────────────────────────────────────────────────

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
MODEL_NOTEBOOK = "GDM_final (2).ipynb"

# Excel inputs — two modes:
#   local    — data.xlsx and salesVariability.xlsx next to app.py (laptop dev only).
#   secrets  — DATA_XLSX + SALES_VAR_XLSX (env or st.secrets); GDM_* names still work as aliases.
#              On Streamlit Cloud, secrets mode is automatic so workbooks need not be on GitHub.
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


def _resolve_data_xlsx_raw():
    """Prefer DATA_XLSX; fall back to GDM_DATA_XLSX."""
    return _resolve_secret_path("DATA_XLSX") or _resolve_secret_path("GDM_DATA_XLSX")


def _resolve_sales_var_xlsx_raw():
    """Prefer SALES_VAR_XLSX; fall back to GDM_SALES_VAR_XLSX."""
    return _resolve_secret_path("SALES_VAR_XLSX") or _resolve_secret_path("GDM_SALES_VAR_XLSX")


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
    if _resolve_data_xlsx_raw() and _resolve_sales_var_xlsx_raw():
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
    raw = _resolve_data_xlsx_raw()
    if not raw:
        return None
    return _normalize_excel_input(raw)


def _sales_var_xlsx_source():
    if _config_mode() == "local":
        return os.path.join(BASE_DIR, "salesVariability.xlsx")
    raw = _resolve_sales_var_xlsx_raw()
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
# LAUNCH_YEARS is now dynamic — set from a sidebar slider in render_sidebar()

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




# ── Data Loading (cached) ───────────────────────────────────────────────────

@st.cache_data(show_spinner="Loading data from Excel files...")
def load_all_data(main_src=None, sales_src=None):
    """Load and parse both Excel files. Returns a dict of all derived tables.

    main_src / sales_src can be:
      - None          → resolved from secrets / local disk (default behaviour)
      - bytes         → content of an uploaded .xlsx file
      - str / path    → file path or URL (existing behaviour)
    """
    if main_src is None:
        main_src = _data_xlsx_source()
    if sales_src is None:
        sales_src = _sales_var_xlsx_source()
    # If bytes were passed (from st.file_uploader), wrap in BytesIO
    if isinstance(main_src, bytes):
        main_src = io.BytesIO(main_src)
    if isinstance(sales_src, bytes):
        sales_src = io.BytesIO(sales_src)

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

    sv_fy_mu = {}
    for r in range(5, 10):
        if pd.isna(sv_raw.iat[r, 0]):
            continue
        arch = str(sv_raw.iat[r, 0]).strip()
        for mat, col in {85: 1, 95: 2, 105: 3, 115: 4}.items():
            if not pd.isna(sv_raw.iat[r, col]):
                sv_fy_mu[(arch, mat)] = float(sv_raw.iat[r, col])

    sv_fy_sig2 = {}
    for r in range(15, 20):
        if pd.isna(sv_raw.iat[r, 0]):
            continue
        arch = str(sv_raw.iat[r, 0]).strip()
        for mat, col in {85: 1, 95: 2, 105: 3, 115: 4}.items():
            if not pd.isna(sv_raw.iat[r, col]):
                sv_fy_sig2[(arch, mat)] = float(sv_raw.iat[r, col])

    sv_gr_mu = {}
    for r in range(5, 25):
        arch = sv_raw.iat[r, 7]
        mat = sv_raw.iat[r, 8]
        if pd.isna(arch) or pd.isna(mat):
            continue
        arch = str(arch).strip()
        mat = int(float(mat))
        for yr in range(2, 11):
            if not pd.isna(sv_raw.iat[r, 9 + (yr - 2)]):
                sv_gr_mu[(arch, mat, yr)] = float(sv_raw.iat[r, 9 + (yr - 2)])

    sv_gr_sig2 = {}
    for r in range(5, 25):
        arch = sv_raw.iat[r, 24]
        mat = sv_raw.iat[r, 25]
        if pd.isna(arch) or pd.isna(mat):
            continue
        arch = str(arch).strip()
        mat = int(float(mat))
        for yr in range(2, 11):
            if not pd.isna(sv_raw.iat[r, 26 + (yr - 2)]):
                sv_gr_sig2[(arch, mat, yr)] = float(sv_raw.iat[r, 26 + (yr - 2)])

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
        return None, None, None, None, ["No valid products selected."], []

    mults = get_yearly_multipliers(strategy, custom_sliders)
    rng = np.random.default_rng(int(seed))
    run_rows_all = []
    product_results = []
    
    cols = [f"Year {i}" for i in range(1, 11)]
    idx_labels = [
        "Carry-in inventory (from prior year)",
        "Carry-in quality loss (10%)",
        "Planned production",
        "Actual production (after yield & conversion)",
        "Production quality loss (2%)",
        "Total saleable inventory",
        "Sales",
        "Remaining inventory [negative = stockout]",
    ]

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
        
        prod_median = np.median(prod_rows, axis=0)
        df_prod = pd.DataFrame(prod_median.T, columns=cols, index=idx_labels)
        
        rem_prod = prod_rows[:, :, 7]
        sales_prod = df_prod.loc["Sales"].astype(float).values
        ay_prod = int(np.where(sales_prod > 0)[0][-1]) if len(np.where(sales_prod > 0)[0]) else 9
        if year_mode != "last_sales":
            ay_prod = custom_year_idx

        prod_summary_df = pd.DataFrame.from_dict({
            f"Selected year (Year {ay_prod + 1})": _row_stats(rem_prod[:, ay_prod]),
            "End of lifecycle (Year 10)": _row_stats(rem_prod[:, 9]),
        }, orient="index")
        
        product_results.append((arch, maturity, df_prod, prod_summary_df))

    run_rows = np.sum(run_rows_all, axis=0)
    median_rows = np.median(run_rows, axis=0)

    lifecycle_df = pd.DataFrame(
        median_rows.T,
        columns=cols,
        index=idx_labels,
    )

    remaining_all = run_rows[:, :, 7]
    ay_idx = determine_analysis_year(lifecycle_df, year_mode, custom_year_idx)

    sel_rem = remaining_all[:, ay_idx]
    end_rem = remaining_all[:, 9]

    selected_stats = _row_stats(sel_rem)
    end_stats = _row_stats(end_rem)

    summary_df = pd.DataFrame.from_dict(
        {
            f"Selected year (Year {ay_idx + 1})": selected_stats,
            "End of lifecycle (Year 10)": end_stats,
        },
        orient="index",
    )

    return lifecycle_df, summary_df, ay_idx, parsed, warnings, product_results


def build_launch_cohorts(launch_plan_df):
    """From launch plan table, create (archetype, maturity, launch_year, n_products).
    Columns are named Y1, Y2, ... and are read dynamically — no hardcoded LAUNCH_YEARS.
    launch_year is 0-based (Y1->0, Y2->1) so results label Year 1 as the first year.
    """
    cohorts = []
    year_cols = sorted(
        [c for c in launch_plan_df.columns
         if isinstance(c, str) and c.startswith("Y") and c[1:].isdigit()],
        key=lambda c: int(c[1:])
    )
    for _, row in launch_plan_df.iterrows():
        arch = str(row.get("Archetype", "")).strip()
        if not arch or arch.lower() == "nan":
            continue
        maturity = int(row["Maturity"])
        for col in year_cols:
            launch_year = int(col[1:]) - 1   # Y1->0, Y2->1, ... (0-based)
            n = int(row.get(col, 0) or 0)
            if n > 0:
                cohorts.append((arch, maturity, launch_year, n))
    return cohorts


def run_multiyear_launch_sim(data, launch_plan_df, iterations, seed, strategy, custom_sliders,
                             use_floor, min_floor, use_max_carry, max_carryover, threshold):
    """Multi-year launch Monte Carlo aggregated over calendar years."""
    cohorts = build_launch_cohorts(launch_plan_df)
    if not cohorts:
        return None, None, None, ["No launches defined. Fill in at least one launch cell > 0."], []

    max_launch = max(c[2] for c in cohorts)
    horizon_years = max_launch + 10
    rng = np.random.default_rng(int(seed))
    mults = get_yearly_multipliers(strategy, custom_sliders)

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

    all_runs = np.zeros((int(iterations), horizon_years, 8), dtype=float)
    unique_products = sorted({(arch, mat) for (arch, mat, _, _) in cohorts})
    product_results = []
    
    year_cols = [f"Year {i}" for i in range(1, horizon_years + 1)]
    idx_labels = [
        "Carry-in inventory (from prior year)",
        "Carry-in quality loss (10%)",
        "Planned production",
        "Actual production (after yield & conversion)",
        "Production quality loss (2%)",
        "Total saleable inventory",
        "Sales",
        "Remaining inventory [negative = stockout]",
    ]

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
        
        prod_median = np.median(prod_all, axis=0)
        df_prod = pd.DataFrame(prod_median.T, columns=year_cols, index=idx_labels)
        
        rem_prod = prod_all[:, :, 7]
        sales_prod = df_prod.loc["Sales"].astype(float).values
        ay_prod = int(np.where(sales_prod > 0)[0][-1]) if len(np.where(sales_prod > 0)[0]) else horizon_years - 1
        
        prod_summary_df = pd.DataFrame.from_dict({
            f"Last sales year (Year {ay_prod + 1})": _row_stats(rem_prod[:, ay_prod]),
            f"End of horizon (Year {horizon_years})": _row_stats(rem_prod[:, -1]),
        }, orient="index")
        
        product_results.append((arch, mat, df_prod, prod_summary_df))

    median_rows = np.median(all_runs, axis=0)
    lifecycle_df = pd.DataFrame(
        median_rows.T,
        columns=year_cols,
        index=idx_labels,
    )

    remaining_all = all_runs[:, :, 7]
    sales_row = lifecycle_df.loc["Sales"].astype(float).values
    s_idx = np.where(sales_row > 0)[0]
    ay_idx = int(s_idx[-1]) if len(s_idx) else horizon_years - 1

    summary_df = pd.DataFrame.from_dict(
        {
            f"Last sales year (Year {ay_idx + 1})": _row_stats(remaining_all[:, ay_idx]),
            f"End of horizon (Year {horizon_years})": _row_stats(remaining_all[:, -1]),
        },
        orient="index",
    )
    return lifecycle_df, summary_df, ay_idx, [], product_results


# ── Streamlit Dashboard ──────────────────────────────────────────────────────

def render_sidebar(data):
    """Render all sidebar controls and return their current values as a dict."""

    st.sidebar.header("Simulation Mode")
    mode_label = st.sidebar.radio(
        "Mode",
        options=["Single-start portfolio", "Multi-year launch cohorts"],
        index=0,
        help="Single-start: one or more products from Year 1. Multi-year: schedule launches across calendar years.",
    )
    mode = "single" if mode_label == "Single-start portfolio" else "multi"

    st.sidebar.header("Product Selection")
    selected_products = st.sidebar.multiselect(
        "Products (Archetype | Maturity)",
        options=data["product_options"],
        default=[data["product_options"][0]] if data["product_options"] else [],
        help="Archetype = seed treatment type. Maturity = crop days rating (85, 95, 105, 115).",
    )

    # ── Multi-year launch year count (only shown in multi mode) ──────────────
    n_launch_years = 6   # default
    if mode == "multi":
        st.sidebar.header("Launch Plan Settings")
        n_launch_years = st.sidebar.slider(
            "Number of launch years",
            min_value=1, max_value=15, value=6, step=1,
            help="Number of launch year columns in the grid. Y1 = first simulation year.",
        )

    st.sidebar.header("Production Strategy")
    strategy_label = st.sidebar.selectbox(
        "Strategy",
        options=list(STRATEGY_OPTIONS.keys()),
        index=0,
        help="How much to produce relative to next year's expected sales. Custom lets you set a multiplier for each year.",
    )
    strategy = STRATEGY_OPTIONS[strategy_label]

    # ── Per-year multipliers: unbounded number inputs (no slider) ─────────────
    custom_multipliers = [1.5] * 10
    if strategy == "custom":
        st.sidebar.markdown(
            "**Per-year production multipliers** "
            "<small style='color:#888'>(any value ≥ 0, including 0 to stop production)</small>",
            unsafe_allow_html=True,
        )
        cols_left, cols_right = st.sidebar.columns(2)
        for i in range(10):
            target = cols_left if i < 5 else cols_right
            custom_multipliers[i] = target.number_input(
                f"Year {i + 1}",
                min_value=0.0,
                value=1.5,
                step=0.1,
                format="%.2f",
                key=f"mult_y{i + 1}",
                help=f"Multiplier for Year {i + 1}. 0 = no production. 1.5 = produce 1.5x next year's sales.",
            )

    st.sidebar.header("Simulation Settings")
    iterations = st.sidebar.slider(
        "Number of simulations", 100, 5000, 1000, step=100,
        help="Number of Monte Carlo runs. More = more stable results but slower. 1,000 recommended.",
    )
    seed = st.sidebar.number_input(
        "Random seed", value=42, step=1,
        help="Fixes the random generator for reproducible results.",
    )

    # ── View mode toggle (restored) ───────────────────────────────────────────
    st.sidebar.header("Results View")
    view_mode = st.sidebar.radio(
        "View mode",
        options=["Table", "Chart"],
        index=0,
        help="Table: lifecycle data and probability summaries. Chart: sales vs production, supply vs demand, waterfall.",
    )
    st.session_state["view_mode"] = view_mode.lower()

    # Analysis Year — single mode only
    year_mode = "last_sales"
    year_mode_label = "Last Year of Sales"
    custom_year_idx = 4
    if mode == "single":
        st.sidebar.header("Analysis Year")
        year_mode_label = st.sidebar.selectbox(
            "Year mode",
            options=["Last Year of Sales", "Choose specific year"],
            help="Last Year of Sales: auto-detects lifecycle end. Choose specific year: manually pick the year.",
        )
        year_mode = "last_sales" if year_mode_label == "Last Year of Sales" else "custom"
        if year_mode == "custom":
            custom_year = st.sidebar.selectbox(
                "Analysis year",
                options=[f"Year {i}" for i in range(1, 11)],
                index=4,
            )
            custom_year_idx = int(custom_year.split()[-1]) - 1

    threshold = st.sidebar.slider(
        "Remaining inventory threshold",
        min_value=0.0, max_value=50000.0, value=0.0, step=500.0,
        help="Reports P(remaining inventory > this value) — the main obsolescence risk metric.",
    )

    st.sidebar.header("Production Constraints")
    use_floor = st.sidebar.checkbox(
        "Enable minimum production floor", value=False,
        help="Ensures planned production never falls below a minimum batch size. Does not force production in zero-production years.",
    )
    min_floor = 0.0
    if use_floor:
        min_floor = st.sidebar.number_input(
            "Min production floor (units)",
            min_value=0.0, value=1000.0, step=100.0, format="%.0f",
            help="Minimum number of units to produce in any year where production is planned.",
        )

    use_max_carry = st.sidebar.checkbox(
        "Enable maximum carryover cap", value=False,
        help="If carry-in inventory meets or exceeds this cap, production is skipped that year.",
    )
    max_carryover = 0.0
    if use_max_carry:
        max_carryover = st.sidebar.number_input(
            "Max carry-in inventory (units)",
            min_value=0.0, value=10000.0, step=500.0, format="%.0f",
            help="Skip production when carry-in inventory reaches this level.",
        )

    return {
        "mode": mode,
        "mode_label": mode_label,
        "products": selected_products,
        "strategy": strategy,
        "strategy_label": strategy_label,
        "custom_sliders": custom_multipliers,
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
        "n_launch_years": n_launch_years,
        "launch_plan_df": None,  # Populated in main area for multi-year mode
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

def render_results(lifecycle_df, summary_df, ay_idx, parsed, warnings, params, product_results=None):
    """Render simulation results in the main area."""

    # Warnings
    for w in (warnings or []):
        st.warning(w)

    if lifecycle_df is None:
        if params["mode"] == "multi":
            st.error("No valid launch cohorts. Please enter at least one launch count (Y1 or later) in the table.")
        else:
            st.error("No valid products selected. Please choose at least one product.")
        return

    # Build year ordering from ay_idx or dataframe columns
    year_order = []
    if isinstance(ay_idx, dict) and ay_idx:
        for y in sorted(ay_idx.keys()):
            year_order.append(f"Y{y}")
    if not year_order and lifecycle_df is not None:
        year_order = [c for c in lifecycle_df.columns if isinstance(c, str) and c.startswith("Y")]

    view_mode = st.session_state.get("view_mode", "table")

    if view_mode == "table":
        if product_results:
            for arch, mat, df_prod, summary_prod in product_results:
                if params.get("mode") == "multi":
                    with st.expander(f"{arch} | Maturity {mat}"):
                        st.dataframe(df_prod.round(1), use_container_width=True)
                        st.dataframe(summary_prod.round(3), use_container_width=True)
                        # AI interpretation for each individual product
                        st.write("DEBUG: calling explain_table on individual product lifecycle")
                        explain_table(df_prod, f"{arch} | Maturity {mat} lifecycle")
                else:
                    st.subheader(f"{arch} | Maturity {mat}")
                    st.dataframe(df_prod.round(1), use_container_width=True)
                    st.dataframe(summary_prod.round(3), use_container_width=True)
                    # AI interpretation for each individual product (single mode)
                    st.write("DEBUG: calling explain_table on individual product lifecycle")
                    explain_table(df_prod, f"{arch} | Maturity {mat} lifecycle")

            if len(product_results) > 1:
                st.subheader("Portfolio Aggregate")
                st.dataframe(lifecycle_df.round(1), use_container_width=True)
                st.write("DEBUG: calling explain_table on portfolio lifecycle")
                explain_table(lifecycle_df, "Portfolio aggregate lifecycle")

                st.subheader("Portfolio Inventory Probability Summary")
                st.dataframe(summary_df.round(3), use_container_width=True)
                st.write("DEBUG: calling explain_table on portfolio summary")
                explain_table(summary_df, "Portfolio aggregate probability summary")
        else:
            st.subheader("Median Lifecycle Across All Simulations")
            st.dataframe(lifecycle_df.round(1), use_container_width=True)
            st.write("DEBUG: calling explain_table on median lifecycle across all runs")
            explain_table(lifecycle_df, "Median lifecycle across all runs")

            st.subheader("Inventory Probability Summary")
            st.dataframe(summary_df.round(3), use_container_width=True)
            st.write("DEBUG: calling explain_table on global probability summary")
            explain_table(summary_df, "Global probability summary")

        st.subheader("Median Remaining Inventory by Year (Aggregate)")
        if "Remaining inventory [negative = stockout]" in lifecycle_df.index:
            remaining_row = lifecycle_df.loc["Remaining inventory [negative = stockout]"]
            chart_df = pd.DataFrame(
                {
                    "Year": year_order,
                    "Remaining Inventory": remaining_row.values[: len(year_order)],
                }
            )
            chart = alt.Chart(chart_df).mark_bar().encode(
                x=alt.X("Year:N", sort=year_order),
                y=alt.Y("Remaining Inventory:Q"),
            )
            st.altair_chart(chart, use_container_width=True)
    else:
        # Chart view — use full _render_chart_view with year ordering
        year_order_for_chart = list(lifecycle_df.columns)
        _render_chart_view(lifecycle_df, year_order_for_chart)

# ── Main ─────────────────────────────────────────────────────────────────────

def main():
    st.set_page_config(
        page_title="Monte Carlo Inventory Planner",
        page_icon="📦",
        layout="wide",
    )

    st.title("Monte Carlo Production & Inventory Planner")
    st.caption("10-year lifecycle simulation with lognormal sales variability, yield & conversion uncertainty")
    if st.button("Test Purdue LLM", key="test_purdue_llm"):
        try:
            payload = {
                "model": MODEL,
                "messages": [
                    {"role": "system", "content": "You are a helpful data analyst."},
                    {"role": "user", "content": "Say 'hello from Purdue'."},
                ],
                "max_tokens": 20,
                "temperature": 0,
            }
            st.write("DEBUG: hitting Purdue endpoint", {"base_url": BASE_URL, "model": MODEL})
            txt = call_llm_api(payload, attempts=1)
            st.success(f"Purdue LLM response: {txt}")
        except Exception as e:
            st.error(f"Purdue test failed: {e}")

    data_src = _data_xlsx_source()
    sv_src = _sales_var_xlsx_source()
    if not _path_readable_for_excel(data_src) or not _path_readable_for_excel(sv_src):
        if _config_mode() == "secrets":
            st.error("Workbook secrets are missing, unreadable, or the URLs could not be loaded.")
            st.markdown(
                "On **Streamlit Cloud**, open your app → **⋮** → **Settings** → **Secrets** and "
                "define **both** keys exactly (names are case-sensitive): **`DATA_XLSX`** and "
                "**`SALES_VAR_XLSX`** (aliases **`GDM_DATA_XLSX`** / **`GDM_SALES_VAR_XLSX`** also work). "
                "Paste a direct **`https://`** link to each file, or "
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

    # ── Optional file uploads (override secrets / local defaults) ───────────
    with st.sidebar.expander("📂 Upload data files (optional)", expanded=False):
        st.markdown(
            "Upload your own `.xlsx` files to override the default data. "
            "Files must follow the same sheet and column structure as the originals. "
            "Leave blank to use the default files loaded from secrets or local disk."
        )
        uploaded_main = st.file_uploader(
            "data.xlsx  (conversion rates, yields, product parameters, median sales)",
            type=["xlsx"],
            key="upload_main",
            help=(
                "Must contain sheets: Conversion rates, Production yields, "
                "Product parameters, Sales volume parameters."
            ),
        )
        uploaded_sv = st.file_uploader(
            "salesVariability.xlsx  (lognormal sales parameters)",
            type=["xlsx"],
            key="upload_sv",
            help=(
                "Must contain sheet: Sales volume parameters with lognormal mu "
                "and sigma² parameters for first-year sales and growth rates."
            ),
        )

    # Resolve sources — uploaded files take priority over defaults
    main_src_arg  = uploaded_main.read() if uploaded_main is not None else None
    sales_src_arg = uploaded_sv.read()   if uploaded_sv   is not None else None

    # Show which files are active
    if uploaded_main or uploaded_sv:
        active_lines = []
        if uploaded_main:
            active_lines.append(f"- **data.xlsx**: uploaded file ({uploaded_main.name})")
        if uploaded_sv:
            active_lines.append(f"- **salesVariability.xlsx**: uploaded file ({uploaded_sv.name})")
        if not uploaded_main:
            active_lines.append("- **data.xlsx**: default (secrets / local disk)")
        if not uploaded_sv:
            active_lines.append("- **salesVariability.xlsx**: default (secrets / local disk)")
        st.info("**Active data sources:**\n" + "\n".join(active_lines))

    data = load_all_data(main_src=main_src_arg, sales_src=sales_src_arg)
    params = render_sidebar(data)

    # ── Main Area Configuration ──────────────────────────────────────────────
    
    # Render main content layout conditionally
    if params["mode"] == "multi":
        st.subheader("Multi-year Launch Cohorts")
        if not params["products"]:
            st.info("Please select one or more products from the sidebar to configure launch cohorts.")
            return
            
        n_ly = params.get("n_launch_years", 6)
        launch_year_cols = [f"Y{y}" for y in range(1, n_ly + 1)]
        st.markdown(
            f"Set the number of products to launch per archetype/maturity per year "
            f"(**Y1 = first simulation year**, Y{n_ly} = last). "
            f"Each product runs for a 10-year lifecycle from its launch year."
        )
        launch_rows = []
        for p in params["products"]:
            a, m_str = p.split(" | ")
            launch_rows.append({
                "Archetype": a.strip(),
                "Maturity": int(m_str.strip()),
                **{col: 0 for col in launch_year_cols},
            })

        params["launch_plan_df"] = st.data_editor(
            pd.DataFrame(launch_rows),
            num_rows="fixed",
            use_container_width=True,
            key="launch_plan_editor",
            column_config={
                "Archetype": st.column_config.TextColumn("Archetype", disabled=True),
                "Maturity":  st.column_config.NumberColumn("Maturity", disabled=True),
                **{
                    col: st.column_config.NumberColumn(
                        col,
                        min_value=0,
                        max_value=99,
                        step=1,
                        help=f"Number of products to launch in {col} (Year {int(col[1:])}).",
                    )
                    for col in launch_year_cols
                },
            },
        )
        
        button_text = "Run Multi-year Simulation"
        run_clicked = st.button(button_text, type="primary")

    else:
        if not params["products"]:
            st.info("Please select one or more products from the sidebar, then run the simulation.")
            return
            
        button_text = "Run Single-start Simulation"
        run_clicked = st.button(button_text, type="primary")

    if run_clicked:
        with st.spinner("Running Monte Carlo simulation..."):
            if params["mode"] == "single":
                lifecycle_df, summary_df, ay_idx, parsed, warnings, product_results = build_lifecycle_sim(
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
                lifecycle_df, summary_df, ay_idx, warnings, product_results = run_multiyear_launch_sim(
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
            "product_results": product_results,
        }

    if "results" in st.session_state:
        r = st.session_state["results"]
        # Clear/Hide the previous results if the user switched modes and hasn't run it yet
        if r["params"]["mode"] != params["mode"]:
            st.info(f"Mode toggled to **{params['mode_label']}**. Click **{button_text}** to update results.")
        else:
            render_results(
                r["lifecycle_df"],
                r["summary_df"],
                r["ay_idx"],
                r["parsed"],
                r["warnings"],
                r["params"],
                r.get("product_results", None),
            )


if __name__ == "__main__":
    main()