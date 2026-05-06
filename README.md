# GDM Decision Simulation Tool

A Streamlit dashboard for running multi-year Monte Carlo simulations of seed production, sales, and inventory. The tool models yield and conversion uncertainty, lifecycle-driven demand, and inventory carryover to quantify stockout risk and excess-inventory risk under different production strategies.

## Project objective

This application was developed as part of the Spring 2026 Industry Practicum with GDM/AgReliant. It is designed to help planners:

- Evaluate production strategies over a 10-year product lifecycle.
- Quantify stockout risk and carryover inventory under historical volatility.
- Compare outcomes across archetype-maturity segments and production strategies.

## Repository structure

- `app.py` - Main Streamlit dashboard application, including the UI, simulation engine, and optional AI interpretation logic.
- `.streamlit/` - Streamlit configuration and local secrets support.
- `validation.py` - Validation and QA utilities for checking simulation outputs and assumptions.
- `requirements.txt` - Python dependencies required to run the app.
- `GDM_Mar17.ipynb` - Early exploratory notebook and prototype work.
- `GDM_final.ipynb` - Final analysis notebook used to design and test the production and inventory logic.
- `GDM_final (2).ipynb` - Alternate export of the final notebook.
- `IP Carryover Model 1.xlsx` - Excel prototype of the carryover model logic.
- `README.md` - Project documentation.

If you rename or remove legacy notebooks, update this section accordingly.

## Data inputs

By default, the app expects the following files in the project root:

- `data.xlsx`
  - Production yields
  - Conversion rates
  - Product parameters
  - Median / average sales parameters

- `salesVariability.xlsx`
  - Lognormal parameters for first-year sales
  - Lognormal parameters for year-over-year growth and decline rates

These files define the distributions used by the Monte Carlo simulation engine.

## Installation

### 1. Create and activate a virtual environment

**Windows (PowerShell)**

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
```

**Windows (Command Prompt)**

```cmd
python -m venv .venv
.venv\Scripts\activate.bat
```

**macOS / Linux**

```bash
python3 -m venv .venv
source .venv/bin/activate
```

### 2. Install dependencies

```bash
pip install -r requirements.txt
```

### 3. Optional: Configure AI interpretation

The app supports optional AI-generated interpretation for simulation outputs and strategy comparison. AI configuration is read at runtime from environment variables or Streamlit secrets.

Supported settings:

- `LLM_API_KEY` - API key for the LLM provider
- `LLM_BASE_URL` - Full API endpoint URL
- `LLM_MODEL` - Model name sent in the request body
- `ENABLE_LLM_EXPLAIN` - Global on/off switch for AI interpretation (`"true"` or `"false"`)
- `LLM_TIMEOUT` - Request timeout in seconds (default: `45`)
- `LLM_DEBUG` - Optional debug flag for showing fallback/debug context (`"true"` or `"false"`)

If AI configuration is missing or invalid, the app will use its built-in fallback interpretation when available instead of failing completely.

#### Local development

For local development, create a file at:

```bash
.streamlit/secrets.toml
```

Example:

```toml
LLM_API_KEY = "your-new-api-key"
LLM_BASE_URL = "https://api.openai.com/v1/chat/completions"
LLM_MODEL = "gpt-4.1"
ENABLE_LLM_EXPLAIN = "true"
LLM_TIMEOUT = "45"
LLM_DEBUG = "false"
```

You can also set the same values as environment variables instead of using `secrets.toml`.

**Important:** do not commit `.streamlit/secrets.toml` to Git.

#### Streamlit Community Cloud

For the deployed app, update the same values in:

**App Settings -> Secrets**

Paste the same block into the Secrets editor:

```toml
LLM_API_KEY = "your-new-api-key"
LLM_BASE_URL = "https://api.openai.com/v1/chat/completions"
LLM_MODEL = "gpt-4.1"
ENABLE_LLM_EXPLAIN = "true"
LLM_TIMEOUT = "45"
LLM_DEBUG = "false"
```

After saving, restart or refresh the app and run a simulation to confirm the new key is working.

#### Rotating or changing the API key

To replace the current LLM key:

1. Create a new API key with your preferred provider or account.
2. Update `LLM_API_KEY` in local `.streamlit/secrets.toml` or in Streamlit Community Cloud Secrets.
3. Test the app by running a simulation and checking that AI interpretation still appears when enabled.
4. Revoke the old key after confirming the new one works.

The sidebar toggle **Enable AI interpretation** controls the feature for the current user session, but backend credentials remain admin-managed through secrets.

### 4. Run the app

```bash
streamlit run app.py
```

The app will open in your browser at `http://localhost:8501`.

If AI interpretation is configured, the app will use the external LLM backend. Otherwise, it will use the built-in fallback interpretation.

## Optional: external venv for local Windows development

If creating `.venv` inside the project fails, for example because of path length, antivirus restrictions, or file-copy errors, you can create an external environment such as `C:\venvs\ip-app`.

### One-time setup

```powershell
New-Item -ItemType Directory -Force -Path C:\venvs

py -m venv C:\venvs\ip-app
C:\venvs\ip-app\Scripts\Activate.ps1
cd "D:\Niharika\Purdue BAIM Courses\Spring 26\Mod 4\IP"
pip install -r requirements.txt
streamlit run app.py
```

Adjust the `cd` path if your project folder is elsewhere.

### Each new terminal session

```powershell
C:\venvs\ip-app\Scripts\Activate.ps1
cd "D:\Niharika\Purdue BAIM Courses\Spring 26\Mod 4\IP"
streamlit run app.py
```

## How the tool is used

1. **Select products** from the sidebar using archetype-maturity combinations.
2. **Choose a production strategy** such as a custom year-by-year multiplier, just-in-time, conservative, or aggressive.
3. **Configure simulation settings** including the number of Monte Carlo iterations and the random seed for reproducibility.
4. **Set analysis parameters** such as year mode, remaining inventory threshold, and production constraints.
5. Click **Run Simulation** to execute the Monte Carlo engine.
6. Review the output tables, summary metrics, risk measures, charts, and optional AI-generated interpretations.

## Output summary

The dashboard can display:

- A year-by-year lifecycle table showing planned production, actual production, planned sales, actual sales, unmet demand, and remaining inventory.
- A summary table with mean and median total sales, mean and median remaining inventory, and risk probabilities.
- Key metric cards and remaining inventory charts.
- A strategy comparison table across fixed strategy tiers using identical assumptions.
- Optional AI-generated or fallback-generated narrative interpretation for lifecycle, summary, product-level, and strategy-comparison outputs.

## Methodology summary

- Yield and conversion variability are estimated from historical production data and modeled as random draws around planned values.
- Lifecycle demand is based on archetype-maturity sales curves, including first-year volume and year-over-year growth or decline factors.
- Inventory evolves year by year with production, realized demand, quality loss, and carryover degradation.
- Monte Carlo simulation is used to generate many possible futures, and the dashboard summarizes these paths into expected values and risk probabilities for planning decisions.

## Notes

- Local mode expects `data.xlsx` and `salesVariability.xlsx` to sit next to `app.py`.
- The app also supports secret-based Excel input configuration for deployment environments.
- If AI interpretation is enabled but the external LLM call fails, the app will fall back to built-in rule-based interpretation logic where supported.
- Strategy comparison uses identically seeded assumptions so the tradeoff across strategy tiers remains directly comparable.

## Troubleshooting

### App starts but AI interpretation does not appear

Check the following:

- `ENABLE_LLM_EXPLAIN` is set to `"true"`.
- `LLM_API_KEY`, `LLM_BASE_URL`, and `LLM_MODEL` are all configured.
- The sidebar toggle **Enable AI interpretation** is turned on.
- Your API endpoint returns an OpenAI-compatible response schema that includes `choices[0].message.content` or `choices[0].text`.

### AI interpretation appears, but it seems repetitive

This can happen if:

- the fallback interpreter is being used,
- the same inputs and random seed are reused,
- or the strategy comparison prompt is highly constrained and produces similar summaries.

### Excel files are not being found

Check that:

- `data.xlsx` and `salesVariability.xlsx` are present in the project root for local mode, or
- the deployment environment is correctly configured to use secret-based workbook input.

## Future improvements

Potential future enhancements include:

- richer scenario management and saved runs,
- improved export options for simulation outputs,
- expanded strategy templates,
- and more transparent AI debug/state reporting in the UI.
