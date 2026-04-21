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
import validation

# ── LLM Helpers ──────────────────────────────────────────────────────────────

def resolve_llm_setting(key, default=None):
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
    return default

def get_llm_config():
    enable_str = resolve_llm_setting("ENABLE_LLM_EXPLAIN", "true")
    debug_str = resolve_llm_setting("LLM_DEBUG", "false")
    return {
        "enable": enable_str.lower() in ("true", "1", "yes"),
        "debug": debug_str.lower() in ("true", "1", "yes"),
        "base_url": resolve_llm_setting("LLM_BASE_URL", "https://genai.rcac.purdue.edu/api/chat/completions"),
        "model": resolve_llm_setting("LLM_MODEL", "llama3.1:latest"),
        "api_key": resolve_llm_setting("LLM_API_KEY", None),
        "timeout": int(resolve_llm_setting("LLM_TIMEOUT", "45")),
    }

def extract_llm_text(data):
    if not isinstance(data, dict):
        raise ValueError("Invalid JSON response schema")
    
    choices = data.get("choices")
    if not choices or not isinstance(choices, list):
        raise ValueError("Response missing 'choices' array")
        
    choice = choices[0]
    
    if "message" in choice and isinstance(choice["message"], dict):
        msg = choice["message"]
        if "content" in msg:
            return msg["content"]
            
    if "text" in choice:
        return choice["text"]
        
    raise ValueError("Could not extract text or message content from response")

def call_llm_api(payload, config, attempts=3):
    if not config["api_key"]:
        raise ValueError("missing API key")
        
    headers = {
        "Authorization": f"Bearer {config['api_key']}",
        "Content-Type": "application/json",
    }

    for attempt in range(attempts):
        try:
            resp = requests.post(config["base_url"], headers=headers, json=payload, timeout=config["timeout"])
            if resp.status_code == 200:
                data = resp.json()
                return extract_llm_text(data)
            
            if resp.status_code == 429 and attempt < attempts - 1:
                time.sleep(2 ** attempt)
                continue
                
            raise RuntimeError(f"HTTP {resp.status_code}: {resp.text}")
            
        except (requests.exceptions.ReadTimeout, requests.exceptions.Timeout):
            if attempt < attempts - 1:
                time.sleep(2 ** attempt)
                continue
            raise RuntimeError(f"request timeout")
        except requests.exceptions.ConnectionError:
            if attempt < attempts - 1:
                time.sleep(2 ** attempt)
                continue
            raise RuntimeError(f"connection error")
        except ValueError as e:
            raise RuntimeError(str(e)) # parsing issues shouldn't retry
            
    raise RuntimeError(f"API call failed after {attempts} attempts")

def build_fallback_explanation(df, context_text=""):
    """Local rule-based fallback interpretation when external API fails."""
    lines = []
    
    if "Strategy" in df.columns:
        try:
            best_sales_idx = df["Mean total actual sales"].idxmax()
            best_sales = df.loc[best_sales_idx, "Strategy"]
            best_sales_val = df.loc[best_sales_idx, "Mean total actual sales"]
            
            best_rem_idx = df["Mean remaining inventory"].idxmin()
            best_rem = df.loc[best_rem_idx, "Strategy"]
            best_rem_val = df.loc[best_rem_idx, "Mean remaining inventory"]
            
            lines.append(f"Across the evaluated strategies, **{best_sales}** generates the highest mean actual sales ({best_sales_val:,.0f} units), prioritizing revenue capture.")
            lines.append(f"Conversely, **{best_rem}** carries the lowest mean remaining inventory ({best_rem_val:,.0f} units), minimizing absolute obsolescence risks.")
            
            if best_sales != best_rem:
                lines.append(f"A structural tradeoff is evident: moving toward **{best_sales}** materially increases revenue probability, but substantially exposes the cycle to expected carryover.")
            else:
                lines.append(f"**{best_sales}** appears broadly dominant, optimizing both total sales bandwidth and residual inventory clearance simultaneously within these parameters.")
                
            max_deplete = df["P(depleted)"].max()
            if max_deplete > 0.1:
                lines.append(f"Aggressive growth parameterizations must be carefully weighed against stock-out consequences, as depletion probabilities fluctuate up to {max_deplete*100:.1f}%.")
            else:
                lines.append("System depletion risks remain functionally suppressed across all strategy matrices, indicating sufficient safety buffers regardless of tuning constraints.")
        except Exception:
            lines.append("Stochastic modeling across the requested strategy permutations reveals distinctive tradeoffs between gross sales capture and terminal inventory carryover.")

    elif "P(depleted)" in df.columns or "Mean remaining inventory" in df.columns:
        # It's a summary table
        # 1. Production/General 
        lines.append(f"Aggregate probabilistic models for {context_text} demonstrate varying degrees of supply constraint and operational buffering across the simulated periods.")
        
        # 2. Sales fulfillment vs Depletion (Inventory)
        max_deplete = 0
        if "P(depleted)" in df.columns:
            max_deplete = df["P(depleted)"].max()
            
        if max_deplete > 0.5:
            lines.append(f"There is a high risk of inventory depletion (peaking at {max_deplete*100:.1f}%), suggesting planned buffers are frequently insufficient against demand volatility.")
        elif max_deplete > 0.1:
            lines.append("A moderate risk of inventory depletion is present, indicating that stochastic demand spikes or yield losses occasionally consume the entire safety stock.")
        else:
            lines.append("Inventory buffers remain highly robust across the entire horizon, effectively minimizing the risk of total depletion.")
            
        # 3. Lifecycle Risk / Horizon (Mean Remaining)
        if "Mean remaining inventory" in df.columns:
            end_mean = df["Mean remaining inventory"].iloc[-1]
            if end_mean > 0:
                lines.append(f"Positive ending inventory (averaging {end_mean:,.0f} units) suggests some residual overbuild or obsolescence risk that should be monitored.")
            else:
                lines.append("End-of-horizon inventory approaches zero, limiting end-of-lifecycle carryover.")
                
    else:
        # It's a lifecycle table
        # 1. Production execution
        planned_prod = 0
        actual_prod = 0
        if "Planned production" in df.index and "Actual production (after yield & conversion)" in df.index:
            planned_prod = df.loc["Planned production"].astype(float).sum()
            actual_prod = df.loc["Actual production (after yield & conversion)"].astype(float).sum()
            
        if actual_prod < planned_prod * 0.98:
            lines.append("Actual production trails planned production, suggesting yield or conversion losses reduce available supply.")
        elif actual_prod > 0:
            lines.append("Actual production tracks planned production closely, indicating operations are broadly meeting targets.")
        else:
            lines.append("Production metrics remain strictly negligible across this specific trajectory.")

        # 2. Sales fulfillment
        planned_sales = 0
        actual_sales = 0
        if "Planned Sales" in df.index and "Actual Sales" in df.index:
            planned_sales = df.loc["Planned Sales"].astype(float).sum()
            actual_sales = df.loc["Actual Sales"].astype(float).sum()
            
        missed = 0
        if "Unmet demand (lost sales)" in df.index:
            missed = df.loc["Unmet demand (lost sales)"].astype(float).sum()
            
        if missed > 0:
            lines.append("Sales appear fully absorbed by available inventory, which may indicate a lean plan with little excess buffer.")
        elif actual_sales >= planned_sales * 0.95:
            lines.append("Actual sales track planned sales closely, indicating demand is broadly being met under the current strategy.")
        else:
            lines.append("Actual sales fall below planned sales, indicating softer demand relative to the forecast.")

        # 3. Inventory / Depletion
        rem_sum = 0
        if "Remaining inventory [0 = depleted]" in df.index:
            rem_arr = df.loc["Remaining inventory [0 = depleted]"].astype(float).values
            rem_sum = sum(rem_arr)
            depleted_years = sum(1 for x in rem_arr if x <= 0.01)
            
            if depleted_years > 2:
                lines.append("Repeated depletion suggests the plan appears to run with limited inventory buffer, increasing exposure to volatility.")
            elif depleted_years > 0:
                lines.append("Periodic inventory depletion occurs, signaling moments where the entire supply buffer was fully drawn down.")
            else:
                lines.append("Remaining inventory tends to carry over across the horizon without sustained full depletion.")

        # 4. End-of-horizon implication
        rem_end = 0
        if "Remaining inventory [0 = depleted]" in df.index:
            rem_end = df.loc["Remaining inventory [0 = depleted]"].astype(float).values[-1]
            if rem_end > (planned_sales * 0.1 if planned_sales else 1000):
                lines.append("Positive ending inventory suggests some residual overbuild or obsolescence risk that should be monitored.")
            elif rem_end <= 0.01:
                lines.append("The product concludes its lifecycle with depletion, limiting end-of-lifecycle carryover.")
            else:
                lines.append("A minor terminal inventory balance securely remains, reflecting a generally well-timed production ramp-down.")
                
    clean_lines = [x.strip() for x in lines if x and str(x).strip()]
    return "\n".join(f"- {x}" for x in clean_lines)

def explain_table(df, context_text="", custom_prompt=None):
    """Builds a JSON payload from a DataFrame and returns a dictionary with text and debug strings."""
    config = get_llm_config()
    
    if not config["enable"]:
        return {"text": None, "_debug_reason": None}

    try:
        # Build compact payload representation of the dataframe
        df_json = df.to_json(orient="split")
        
        if custom_prompt:
            prompt = custom_prompt
        else:
            prompt = f"""
Provide exactly 4 markdown bullet points interpreting this Monte Carlo simulation table for a business audience.

Context: {context_text}
Table data: {df_json}

IMPORTANT:
- All quantities are in bushels/units of seed, not dollars.
- Do not use dollar signs or currency symbols.
- Each bullet should be 1 to 2 full sentences.
- Output bullet points only, with each bullet on its own line.
- Do not write a paragraph, title, intro, or conclusion.
- Planned Sales drives deterministic production.
- Actual Sales equals realized demand capped at available inventory.

When supported by the data, cover:
1. planned vs actual production,
2. sales fulfillment vs unmet demand,
3. remaining inventory and depletion risk,
4. the overall lifecycle or end-of-horizon implication.
"""

        payload = {
            "model": config["model"],
            "messages": [
                {"role": "system", "content": "You are a helpful data analyst."},
                {"role": "user", "content": prompt},
            ],
            "max_tokens": 350,
            "temperature": 0.4,
        }
        
    except Exception as e:
        return {
            "text": build_fallback_explanation(df, context_text),
            "_debug_reason": f"payload generation error: {str(e)}"
        }

    try:
        explanation = call_llm_api(payload, config, attempts=3)
        return {"text": explanation, "_debug_reason": None}
    except Exception as e:
        reason = str(e)
        return {
            "text": build_fallback_explanation(df, context_text),
            "_debug_reason": reason
        }

def generate_all_explanations(lifecycle_df, summary_df, product_results):
    """Generates all AI explanations after a successful simulation run."""
    explanations = {}
    config = get_llm_config()
    
    if not config["enable"]:
        return explanations
    
    with st.spinner("Generating AI explanations..."):
        if product_results:
            explanations["product_results"] = {}
            for arch, mat, df_prod, summary_prod in product_results:
                key = f"{arch} | Maturity {mat}"
                explanations["product_results"][key] = explain_table(df_prod, f"{key} lifecycle")
            
            if len(product_results) > 1:
                explanations["lifecycle"] = explain_table(lifecycle_df, "Portfolio aggregate lifecycle")
                explanations["summary"] = explain_table(summary_df, "Portfolio remaining inventory summary")
        else:
            explanations["lifecycle"] = explain_table(lifecycle_df, "Sample lifecycle track")
            explanations["summary"] = explain_table(summary_df, "Global remaining inventory summary")
            
    return explanations

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
    with pd.ExcelFile(main_src) as main_xls, pd.ExcelFile(sales_src) as sales_xls:
        main_val = validation.validate_main_workbook_structure(main_xls)
        sales_val = validation.validate_sales_variability_workbook_structure(sales_xls)
        
        if not main_val["ok"] or not sales_val["ok"]:
            errs = validation.format_validation_errors(main_val, sales_val)
            raise validation.WorkbookValidationError(errs)

        conv_sheet = validation.resolve_sheet_name(main_xls, ["conversion"])
        yield_sheet = validation.resolve_sheet_name(main_xls, ["yield"])
        params_sheet = validation.resolve_sheet_name(main_xls, ["product param"])
        sales_sheet = validation.resolve_sheet_name(main_xls, ["sales", "vol"])

        conv_tab = pd.read_excel(main_xls, sheet_name=conv_sheet)
        yield_tab = pd.read_excel(main_xls, sheet_name=yield_sheet)
        params_tab = pd.read_excel(main_xls, sheet_name=params_sheet)
        sales_raw = pd.read_excel(main_xls, sheet_name=sales_sheet, header=None)
        
        sv_fy_mu, sv_fy_sig2, sv_gr_mu, sv_gr_sig2 = validation.extract_sales_variability(sales_xls)

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
    # Automatically extracted via anchor parsing in validation.py
    # and securely assigned in the `with` block above properly.
    sv_gr_sig2_dict = sv_gr_sig2

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


def simulate_one_run(data, archetype, maturity, rng, mults, use_floor, min_floor,
                     use_max_carry, max_carryover, strategy="custom"):
    """
    One Monte Carlo lifecycle:
      data          -- parsed from excel
      archetype     -- product archetype
      maturity      -- product maturity
      rng           -- numpy random generator
      mults         -- 10 production multipliers
...
    """
    carryover = 0.0
    rows = []
    missed_sales = []
    
    actual_sales_list = []
    planned_sales_list = []
    prior_demand = 0.0

    for yr in range(10):
        # 1) Determine planned sales and realized demand
        if yr == 0:
            fy_mu = data["sv_fy_mu"].get((archetype, maturity))
            fy_sig2 = data["sv_fy_sig2"].get((archetype, maturity))
            if fy_mu is not None and fy_sig2 is not None:
                ev_y1 = float(np.exp(fy_mu + fy_sig2 / 2.0))
                fy_sigma = float(np.sqrt(max(fy_sig2, 1e-10)))
                y1_demand = float(rng.lognormal(fy_mu, fy_sigma))
                expected_sales_target = ev_y1
                prior_demand = y1_demand
                realized_demand = prior_demand
            else:
                val = get_median_sales(data, archetype, maturity) or 0.0
                expected_sales_target = val
                prior_demand = val
                realized_demand = prior_demand
        else:
            prior_expected = planned_sales_list[-1]
            gr_mu = data["sv_gr_mu"].get((archetype, maturity, yr + 1))
            gr_sig2 = data["sv_gr_sig2"].get((archetype, maturity, yr + 1))
            if gr_mu is not None and gr_sig2 is not None:
                ev_growth = float(np.exp(gr_mu + gr_sig2 / 2.0))
                gr_sigma = float(np.sqrt(max(gr_sig2, 1e-10)))
                growth_draw = float(rng.lognormal(gr_mu, gr_sigma))
            else:
                yoy = get_yoy_rates(data, archetype, maturity)
                if yoy is not None:
                    ev_growth = 1.0 + yoy[yr - 1]
                    growth_draw = ev_growth
                else:
                    ev_growth = 0.0
                    growth_draw = 0.0
            
            expected_sales_target = prior_expected * ev_growth
            prior_demand = prior_demand * growth_draw
            realized_demand = prior_demand

        planned_sales_list.append(expected_sales_target)

        # 2) Planned production
        planned_prod = mults[yr] * expected_sales_target

        if use_max_carry and max_carryover > 0.0 and carryover >= max_carryover:
            planned_prod = 0.0

        if use_floor and planned_prod > 0.0:
            planned_prod = max(planned_prod, min_floor)

        # 3) Actual production
        y_draw, c_draw = sample_yield_conv_normal(data, archetype, rng)
        new_prod = planned_prod * y_draw * c_draw

        prod_loss = new_prod * 0.02
        carry_loss = carryover * 0.10

        # 4) Total saleable inventory
        total_saleable = (carryover - carry_loss) + (new_prod - prod_loss)

        # 5) Actual sales & inventory constraint
        # Actual sales are realized demand capped at total saleable inventory (not capped by planned sales).
        actual_sales_yr = min(realized_demand, total_saleable)
        actual_sales_list.append(actual_sales_yr)

        remaining = total_saleable - actual_sales_yr

        missed = max(0.0, realized_demand - total_saleable)
        missed_sales.append(missed)

        rows.append([
            carryover,
            -carry_loss,
            planned_prod,
            new_prod,
            -prod_loss,
            total_saleable,
            expected_sales_target,
            actual_sales_yr,
            remaining,
            missed,
        ])
        carryover = remaining

    return np.array(rows, dtype=float), np.array(missed_sales, dtype=float)

def compute_summary_metrics(rem_arr, total_sales_arr, threshold):
    """Shared summary metric helper for single and multi-year setups."""
    thr = float(threshold)
    stats = {}
    if total_sales_arr is not None:
        stats["Mean total actual sales"] = float(total_sales_arr.mean())
        stats["Median total actual sales"] = float(np.median(total_sales_arr))
        
    stats["Mean remaining inventory"] = float(rem_arr.mean())
    stats["Median remaining inventory"] = float(np.median(rem_arr))
    stats[f"P(remaining > {thr:.0f})"] = float((rem_arr > thr).mean())
    stats["P(depleted)"] = float((rem_arr <= 0).mean())
    
    return stats



def determine_analysis_year(lifecycle_df, year_mode, custom_year_idx):
    """Resolve year-mode selection to a 0-based year index."""
    sales_row = lifecycle_df.loc["Actual Sales"].astype(float).values
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
        "Planned Sales",
        "Actual Sales",
        "Remaining inventory [0 = depleted]",
        "Unmet demand (lost sales)",
    ]

    thr = float(threshold)

    for arch, maturity in parsed:
        prod_rows = []
        for _ in range(int(iterations)):
            rows, _ = simulate_one_run(
                data, arch, maturity, rng, mults,
                use_floor, min_floor, use_max_carry, max_carryover, strategy
            )
            prod_rows.append(rows)
        prod_rows = np.stack(prod_rows, axis=0)
        run_rows_all.append(prod_rows)
        
        prod_median = np.median(prod_rows, axis=0)
        df_prod = pd.DataFrame(prod_median.T, columns=cols, index=idx_labels)
        
        rem_prod = prod_rows[:, :, 8]
        miss_prod = prod_rows[:, :, 9]
        sales_prod = df_prod.loc["Actual Sales"].astype(float).values
        ay_prod = int(np.where(sales_prod > 0)[0][-1]) if len(np.where(sales_prod > 0)[0]) else 9
        if year_mode != "last_sales":
            ay_prod = custom_year_idx

        total_sales_prod = prod_rows[:, :, 7].sum(axis=1)
        prod_summary_df = pd.DataFrame.from_dict({
            f"Selected year (Year {ay_prod + 1})": compute_summary_metrics(rem_prod[:, ay_prod], total_sales_prod, threshold),
            "End of lifecycle (Year 10)": compute_summary_metrics(rem_prod[:, 9], total_sales_prod, threshold),
        }, orient="index")
        
        product_results.append((arch, maturity, df_prod, prod_summary_df))

    run_rows = np.sum(run_rows_all, axis=0)
    median_rows = np.median(run_rows, axis=0)

    lifecycle_df = pd.DataFrame(
        median_rows.T,
        columns=cols,
        index=idx_labels,
    )

    remaining_all = run_rows[:, :, 8]
    missed_all = run_rows[:, :, 9]
    ay_idx = determine_analysis_year(lifecycle_df, year_mode, custom_year_idx)

    sel_rem = remaining_all[:, ay_idx]
    sel_miss = missed_all[:, ay_idx]
    end_rem = remaining_all[:, 9]
    end_miss = missed_all[:, 9]
    total_sales_all = run_rows[:, :, 7].sum(axis=1)

    selected_stats = compute_summary_metrics(sel_rem, total_sales_all, threshold)
    end_stats = compute_summary_metrics(end_rem, total_sales_all, threshold)

    summary_df = pd.DataFrame.from_dict(
        {
            f"Selected year (Year {ay_idx + 1})": selected_stats,
            "End of lifecycle (Year 10)": end_stats,
        },
        orient="index",
    )

    return lifecycle_df, summary_df, ay_idx, parsed, warnings, product_results, run_rows


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

    all_runs = np.zeros((int(iterations), horizon_years, 10), dtype=float)
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
        "Planned Sales",
        "Actual Sales",
        "Remaining inventory [0 = depleted]",
        "Unmet demand (lost sales)",
    ]

    for arch, mat in unique_products:
        base_runs = []
        for _ in range(int(iterations)):
            rows, _ = simulate_one_run(
                data, arch, mat, rng, mults,
                use_floor, min_floor, use_max_carry, max_carryover, strategy
            )
            base_runs.append(rows)
        base_runs = np.stack(base_runs, axis=0)

        prod_all = np.zeros((int(iterations), horizon_years, 10), dtype=float)
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
        
        rem_prod = prod_all[:, :, 8]
        miss_prod = prod_all[:, :, 9]
        sales_prod = df_prod.loc["Actual Sales"].astype(float).values
        ay_prod = int(np.where(sales_prod > 0)[0][-1]) if len(np.where(sales_prod > 0)[0]) else horizon_years - 1
        
        total_sales_prod = prod_all[:, :, 7].sum(axis=1)
        prod_summary_df = pd.DataFrame.from_dict({
            f"Last sales year (Year {ay_prod + 1})": compute_summary_metrics(rem_prod[:, ay_prod], total_sales_prod, threshold),
            f"End of horizon (Year {horizon_years})": compute_summary_metrics(rem_prod[:, -1], total_sales_prod, threshold),
        }, orient="index")
        
        product_results.append((arch, mat, df_prod, prod_summary_df))

    median_rows = np.median(all_runs, axis=0)
    lifecycle_df = pd.DataFrame(
        median_rows.T,
        columns=year_cols,
        index=idx_labels,
    )

    remaining_all = all_runs[:, :, 8]
    missed_all = all_runs[:, :, 9]
    sales_row = lifecycle_df.loc["Actual Sales"].astype(float).values
    s_idx = np.where(sales_row > 0)[0]
    ay_idx = int(s_idx[-1]) if len(s_idx) else horizon_years - 1

    total_sales_all = all_runs[:, :, 7].sum(axis=1)
    summary_df = pd.DataFrame.from_dict(
        {
            f"Last sales year (Year {ay_idx + 1})": compute_summary_metrics(remaining_all[:, ay_idx], total_sales_all, threshold),
            f"End of horizon (Year {horizon_years})": compute_summary_metrics(remaining_all[:, -1], total_sales_all, threshold),
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
    
    if "selected_products" in st.session_state:
        st.session_state["selected_products"] = [
            p for p in st.session_state["selected_products"] 
            if p in data["product_options"]
        ]

    selected_products = st.sidebar.multiselect(
        "Products (Archetype | Maturity)",
        options=data["product_options"],
        default=[data["product_options"][0]] if data["product_options"] else [],
        key="selected_products",
        help="Archetype = seed trait type. Maturity = relative maturity rating (85, 95, 105, 115).",
    )

    # ── Multi-year launch year count (only shown in multi mode) ──────────────
    n_launch_years = 6   # default
    if mode == "multi":
        n_launch_years = st.sidebar.number_input(
            "Number of launch years",
            min_value=1, value=6, step=1,
            help="Type the number of calendar launch years to include in the launch plan. Y1 is the first simulation year.",
        )

    st.sidebar.header("Production Strategy")
    strategy_label = st.sidebar.selectbox(
        "Strategy",
        options=list(STRATEGY_OPTIONS.keys()),
        index=list(STRATEGY_OPTIONS.values()).index("jit"),
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
        "Planned Sales": lifecycle_df.loc["Planned Sales"].values,
        "Actual Sales": lifecycle_df.loc["Actual Sales"].values,
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
            domain=["Planned Production", "Actual Production", "Planned Sales", "Actual Sales"],
            range=["#4c78a8", "#72b7b2", "#ffb570", "#f58518"],
        )),
        xOffset="Metric:N",
    )
    st.altair_chart(sp_chart, use_container_width=True)

    # ── 2. Total Saleable Inventory vs Sales (overlay line) ──────────────
    st.subheader("Total Saleable Inventory vs Sales")
    line_metrics = {
        "Total Saleable Inventory": lifecycle_df.loc["Total saleable inventory"].values,
        "Planned Sales": lifecycle_df.loc["Planned Sales"].values,
        "Actual Sales": lifecycle_df.loc["Actual Sales"].values,
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
            domain=["Total Saleable Inventory", "Planned Sales", "Actual Sales"],
            range=["#4c78a8", "#ffb570", "#f58518"],
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
        ("Actual Sales",      float(-col_data.loc["Actual Sales"])),
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

def render_explanation(explain_obj):
    if not isinstance(explain_obj, dict):
        if isinstance(explain_obj, str): # Legacy catch
            st.markdown(f"**AI interpretation:**\n{explain_obj}")
        return
    if explain_obj.get("text"):
        st.markdown(f"**AI interpretation:**\n{explain_obj['text']}")
    config = get_llm_config()
    if config["debug"] and explain_obj.get("_debug_reason"):
        st.caption(f"Debug: AI fallback used. Reason: {explain_obj['_debug_reason']}")

def render_results(lifecycle_df, summary_df, ay_idx, parsed, warnings, params, product_results=None, explanations=None, run_rows=None):
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

    tab1, tab2 = st.tabs(["Table", "Chart"])
    
    with tab1:
        if product_results:
            for arch, mat, df_prod, summary_prod in product_results:
                key = f"{arch} | Maturity {mat}"
                if params.get("mode") == "multi":
                    with st.expander(key, expanded=True):
                        st.dataframe(df_prod.drop(["Unmet demand (lost sales)"], errors="ignore").round(1), use_container_width=True)
                        st.dataframe(summary_prod.round(3), use_container_width=True)
                        # AI interpretation for each individual product
                        if explanations and key in explanations.get("product_results", {}):
                            render_explanation(explanations['product_results'][key])
                else:
                    st.subheader(key)
                    st.dataframe(df_prod.drop(["Unmet demand (lost sales)"], errors="ignore").round(1), use_container_width=True)
                    st.dataframe(summary_prod.round(3), use_container_width=True)
                    # AI interpretation for each individual product (single mode)
                    if explanations and key in explanations.get("product_results", {}):
                        render_explanation(explanations['product_results'][key])

            if len(product_results) > 1:
                st.subheader("Portfolio Aggregate")
                st.dataframe(lifecycle_df.drop(["Unmet demand (lost sales)"], errors="ignore").round(1), use_container_width=True)
                if explanations and "lifecycle" in explanations:
                    render_explanation(explanations['lifecycle'])

                st.subheader("Portfolio Remaining Inventory Summary")
                st.dataframe(summary_df.round(3), use_container_width=True)
                if explanations and "summary" in explanations:
                    render_explanation(explanations['summary'])
        else:
            st.subheader("Median Lifecycle Track (Across All Runs)")
            st.dataframe(lifecycle_df.drop(["Unmet demand (lost sales)"], errors="ignore").round(1), use_container_width=True)
            if explanations and "lifecycle" in explanations:
                render_explanation(explanations['lifecycle'])

            st.subheader("Remaining Inventory Summary")
            st.dataframe(summary_df.round(3), use_container_width=True)
            if explanations and "summary" in explanations:
                render_explanation(explanations['summary'])
                


        st.subheader("Remaining Inventory by Year")
        if "Remaining inventory [0 = depleted]" in lifecycle_df.index:
            remaining_row = lifecycle_df.loc["Remaining inventory [0 = depleted]"]
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
            
        if params.get("mode") == "single" and run_rows is not None:
            st.divider()
            with st.expander("Advanced Debug: Single-Run Trace Explorer"):
                st.markdown("Examine a specific Monte Carlo iteration to explicitly see the relationship between Planned Sales, Actual Sales, and Inventory limits before median-smoothing.")
                
                found_case1 = -1
                found_case2 = -1
                found_case3 = -1
                
                iterations = run_rows.shape[0]
                for i in range(iterations):
                    arr = run_rows[i]
                    planned = arr[:, 6]
                    actual = arr[:, 7]
                    missed = arr[:, 9]
                    supply = arr[:, 5]
                    
                    if found_case1 == -1 and np.any(actual <= planned):
                        found_case1 = i
                    if found_case2 == -1 and np.any((actual > planned) & (actual < supply) & (missed == 0)):
                        found_case2 = i
                    if found_case3 == -1 and np.any(missed > 0):
                        found_case3 = i
                        
                    if found_case1 >= 0 and found_case2 >= 0 and found_case3 >= 0:
                        break
                        
                opts = {"Run 0 (Default)": 0}
                if found_case1 >= 0: opts["First run showing Case 1 (Actual < Planned)"] = found_case1
                if found_case2 >= 0: opts["First run showing Case 2 (Demand beats plan, fully met)"] = found_case2
                if found_case3 >= 0: opts["First run showing Case 3 (Supply caps demand)"] = found_case3
                opts["Custom Run Index"] = "custom"
                
                chk = st.selectbox("Select trace:", list(opts.keys()))
                if chk == "Custom Run Index":
                    sel_idx = st.number_input("Enter Run Index", min_value=0, max_value=iterations-1, value=0, step=1)
                else:
                    sel_idx = opts[chk]
                
                tgt_arr = run_rows[sel_idx]
                regimes = []
                for yr in range(10):
                    p = tgt_arr[yr, 6]
                    a = tgt_arr[yr, 7]
                    m = tgt_arr[yr, 9]
                    
                    if m > 0:
                        regimes.append("Case 3: Supply Caps Demand")
                    elif a > p and a < tgt_arr[yr, 5] and m == 0:
                        regimes.append("Case 2: Planned < Actual < Supply")
                    else:
                        regimes.append("Case 1: Actual <= Planned")
                        
                trace_df = pd.DataFrame({
                    "Planned Sales": tgt_arr[:, 6],
                    "Actual Sales": tgt_arr[:, 7],
                    "Total saleable inventory": tgt_arr[:, 5],
                    "Unmet demand (lost sales)": tgt_arr[:, 9],
                    "Regime": regimes
                }, index=[f"Year {i+1}" for i in range(10)])
                
                # Format to 1 decimal place, ignoring the strings
                styled_df = trace_df.T.copy()
                for col in styled_df.columns:
                    styled_df[col] = styled_df[col].apply(lambda x: f"{x:.1f}" if isinstance(x, (int, float)) else x)

                st.dataframe(styled_df, use_container_width=True)

                st.caption("")
                st.markdown("##### Reproduce This Trace")
                st.markdown("Copy or download these exact parameters to perfectly recreate this specific simulation path.")
                
                repro_dict = {
                    "mode": params.get("mode"),
                    "products": params.get("products", []),
                    "seed": params.get("seed"),
                    "iterations": params.get("iterations"),
                    "strategy": params.get("strategy"),
                    "custom_sliders": list(params.get("custom_sliders", [])),
                    "use_floor": params.get("use_floor"),
                    "min_floor": params.get("min_floor"),
                    "use_max_carry": params.get("use_max_carry"),
                    "max_carryover": params.get("max_carryover"),
                    "threshold": params.get("threshold"),
                    "year_mode": params.get("year_mode"),
                    "custom_year_idx": params.get("custom_year_idx"),
                    "trace_label": chk,
                    "run_index": int(sel_idx)
                }
                
                repro_json = json.dumps(repro_dict, indent=2)
                
                st.code(repro_json, language="json")
                st.download_button(
                    label="Download Reproducibility JSON",
                    data=repro_json,
                    file_name=f"trace_run_{sel_idx}.json",
                    mime="application/json",
                    key="repro_dl"
                )



    with tab2:
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

    st.markdown(
        """
        <style>
        section[data-testid="stSidebar"] {
            min-width: 290px;
            max-width: 330px;
        }
        </style>
        """,
        unsafe_allow_html=True
    )

    st.title("Monte Carlo Production & Inventory Planner")
    st.caption("10-year lifecycle simulation with lognormal sales variability, yield & conversion uncertainty")


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
    with st.expander("📂 Upload data files (optional)", expanded=False):
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

    try:
        data = load_all_data(main_src=main_src_arg, sales_src=sales_src_arg)
        st.success("Validation Summary: Workbooks loaded successfully.")
    except validation.WorkbookValidationError as e:
        st.error("We couldn't use the uploaded workbooks because they do not match the expected GDM format.")
        for err_msg in e.validation_results:
            st.warning(err_msg)
        st.info("Please upload workbooks that match the GDM template.")
        st.stop()
    except Exception as e:
        st.error(f"An unexpected error occurred while parsing the workbooks: {str(e)}")
        st.stop()
    params = render_sidebar(data)

    with st.expander("🛠 Internal Testing: Case 2 Grid Search (Temporary QA)", expanded=False):
        st.markdown("Search a batch of seeds for combinations where the **first main table** strictly exhibits Case 2 behavior (Actual Sales > Planned Sales AND Unmet demand == 0).")
        col1, col2, col3 = st.columns(3)
        with col1:
            qa_start = st.number_input("Start Seed", value=1)
        with col2:
            qa_end = st.number_input("End Seed", value=100)
        with col3:
            qa_iter = st.number_input("Fixed Iterations", value=1000)
            
        if st.button("Run QA Search Grid"):
            if not params.get("products"):
                st.warning("Please select at least one product in the sidebar.")
            else:
                with st.spinner("Searching seed space..."):
                    qa_results = []
                    for sd in range(qa_start, qa_end + 1):
                        lifecycle_df, _, _, _, _, _, _ = build_lifecycle_sim(
                            data,
                            params["products"],
                            qa_iter,
                            sd,
                            params["strategy"],
                            params["custom_sliders"],
                            params["use_floor"],
                            params["min_floor"],
                            params["use_max_carry"],
                            params["max_carryover"],
                            params["year_mode"],
                            params.get("custom_year_idx", 9),
                            params["threshold"],
                        )
                        planned = lifecycle_df.loc["Planned Sales"].astype(float).values
                        actual = lifecycle_df.loc["Actual Sales"].astype(float).values
                        unmet = lifecycle_df.loc["Unmet demand (lost sales)"].astype(float).values
                        
                        match_yrs = []
                        for idx, (p, a, u) in enumerate(zip(planned, actual, unmet)):
                            if a > p and abs(u) < 1e-9:
                                match_yrs.append(idx + 1)
                                
                        if match_yrs:
                            qa_results.append({
                                "Product(s)": ", ".join(params["products"]),
                                "Strategy": params["strategy"],
                                "Seed": sd,
                                "Iterations": qa_iter,
                                "Floor On": params["use_floor"],
                                "Floor Min": params["min_floor"],
                                "Max Carry On": params["use_max_carry"],
                                "Max Carry": params["max_carryover"],
                                "Matches (Years)": str(match_yrs)
                            })
                            
                    if qa_results:
                        st.success(f"Found {len(qa_results)} matching seeds!")
                        st.dataframe(pd.DataFrame(qa_results), use_container_width=True)
                    else:
                        st.info("No Case 2 matches found in this grid search.")

    # ── Main Area Configuration ──────────────────────────────────────────────
    
    # Render main content layout conditionally
    comp_clicked = False
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

        with st.expander("Configure multi-year launch plan", expanded=True):
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
        st.session_state["comparison_df"] = None
        st.session_state["comparison_explanation"] = None
        with st.spinner("Running Monte Carlo simulation..."):
            if params["mode"] == "single":
                lifecycle_df, summary_df, ay_idx, parsed, warnings, product_results, run_rows = build_lifecycle_sim(
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
                run_rows = None
        
        st.session_state["results"] = {
            "lifecycle_df": lifecycle_df,
            "summary_df": summary_df,
            "ay_idx": ay_idx,
            "parsed": parsed,
            "warnings": warnings,
            "params": params,
            "product_results": product_results,
            "run_rows": run_rows,
            "explanations": {},
            "ai_pending": True,
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
                r.get("explanations", None),
                r.get("run_rows", None),
            )
            
            if r.get("ai_pending", False):
                with st.spinner("Finalizing AI interpretation..."):
                    time.sleep(2)
                    explanations = generate_all_explanations(r["lifecycle_df"], r["summary_df"], r.get("product_results", None))
                    r["explanations"] = explanations
                    r["ai_pending"] = False
                st.rerun()

            st.divider()
            if "comparison_df" not in st.session_state or st.session_state["comparison_df"] is None:
                if st.button("Compare Strategies", type="secondary", help="Generate a strategy tradeoff summary table using identically seeded inputs."):
                    with st.spinner("Running strategy comparison..."):
                        comp_rows = []
                        for s_name, s_id in [("Just-in-time 1.0x", "jit"), ("Conservative 1.2x", "cons"), ("Aggressive 2.0x", "aggr")]:
                            if params["mode"] == "single":
                                _, summary_df, ay_idx, _, _, _, _ = build_lifecycle_sim(
                                    data, params["products"], params["iterations"], params["seed"],
                                    s_id, params["custom_sliders"], params["use_floor"], params["min_floor"],
                                    params["use_max_carry"], params["max_carryover"], params["year_mode"],
                                    params["custom_year_idx"], params["threshold"]
                                )
                                sel_row_name = f"Selected year (Year {ay_idx + 1})"
                            else:
                                _, summary_df, ay_idx, _, _ = run_multiyear_launch_sim(
                                    data, params["launch_plan_df"], params["iterations"], params["seed"],
                                    s_id, params["custom_sliders"], params["use_floor"], params["min_floor"],
                                    params["use_max_carry"], params["max_carryover"], params["threshold"]
                                )
                                sel_row_name = f"Last sales year (Year {ay_idx + 1})"
                            
                            row_data = {
                                "Strategy": s_name,
                                "Mean total actual sales": summary_df.loc[sel_row_name, "Mean total actual sales"],
                                "Median total actual sales": summary_df.loc[sel_row_name, "Median total actual sales"],
                                "Mean remaining inventory": summary_df.loc[sel_row_name, "Mean remaining inventory"],
                                "Median remaining inventory": summary_df.loc[sel_row_name, "Median remaining inventory"],
                                f"P(remaining > {params['threshold']:.0f})": summary_df.loc[sel_row_name, f"P(remaining > {params['threshold']:.0f})"],
                                "P(depleted)": summary_df.loc[sel_row_name, "P(depleted)"],
                            }
                            comp_rows.append(row_data)
                        comp_df = pd.DataFrame(comp_rows)
                        st.session_state["comparison_df"] = comp_df
                        
                        conf = get_llm_config()
                        if conf.get("enable", True):
                            best_sales = comp_df.loc[comp_df["Mean total actual sales"].idxmax(), "Strategy"]
                            best_sales_val = comp_df["Mean total actual sales"].max()
                            
                            best_rem = comp_df.loc[comp_df["Mean remaining inventory"].idxmin(), "Strategy"]
                            best_rem_val = comp_df["Mean remaining inventory"].min()
                            
                            custom_prompt = f"""
Provide exactly 4 markdown bullet points explicitly comparing the strategies in this Monte Carlo tradeoff table.
Context: Strategy tradeoff comparison focusing on actual sales versus depletion/obsolescence risk.
Table data: {comp_df.to_json(orient='split')}

IMPORTANT:
- Ensure each bullet explicitly names specific strategies (e.g. 'Just-in-time 1.0x', 'Conservative 1.2x', 'Aggressive 2.0x').
- Discuss how they differ on 'Mean total actual sales' and 'Mean remaining inventory' or 'P(depleted)'.
- Explicitly note that '{best_sales}' has the highest mean actual sales ({best_sales_val:,.0f}).
- Explicitly note that '{best_rem}' has the lowest mean remaining inventory ({best_rem_val:,.0f}).
- Mention which strategy appears to offer the most balanced tradeoff between high sales and low obsolescence risk.
- Do not use generic single-strategy depletion language. You must compare them against each other.
- Output bullet points only, with each bullet on its own line.
- Do not use dollar signs or currency symbols.
"""
                            st.session_state["comparison_explanation"] = explain_table(comp_df, custom_prompt=custom_prompt)
                    st.rerun()
            else:
                st.subheader("Strategy Comparison")
                st.markdown("Tradeoff analysis across fixed strategy tiers using identical assumptions.")
                comp_df = st.session_state["comparison_df"]
                st.dataframe(comp_df, use_container_width=True)
                csv = comp_df.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="Download Comparison as CSV",
                    data=csv,
                    file_name='strategy_comparison.csv',
                    mime='text/csv',
                )
                
                exp = st.session_state.get("comparison_explanation")
                if exp:
                    render_explanation(exp)

if __name__ == "__main__":
    main()