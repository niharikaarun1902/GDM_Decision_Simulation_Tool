# Monte Carlo Production \& Inventory Planner 

A Streamlit dashboard for running 10-year Monte Carlo simulations of seed production and inventory across multiple products. Uses lognormal sales variability, yield \& conversion uncertainty, and portfolio-level aggregation.

## Prerequisites

- Python 3.10 or later
- The following data files in the project root:
    - `data.xlsx` (conversion rates, production yields, product parameters, median sales)
    - `salesVariability.xlsx` (lognormal parameters for first-year sales and growth rates)


## Setup

### 1. Create and activate the virtual environment

**Windows (PowerShell):**

```powershell
python -m venv .venv
.venv\\Scripts\\Activate.ps1
```

**Windows (Command Prompt):**

```cmd
python -m venv .venv
.venv\\Scripts\\activate.bat
```

**macOS / Linux:**

```bash
python3 -m venv .venv
source .venv/bin/activate
```


### 2. Install dependencies

```
pip install -r requirements.txt
```


### 3. Optional: Configure AI interpretation

The app supports optional AI-generated interpretation for simulation outputs and strategy comparison. AI configuration is read from environment variables or Streamlit secrets at runtime. The relevant settings are:

- `LLM_API_KEY` – API key for the LLM provider
- `LLM_BASE_URL` – full API endpoint URL
- `LLM_MODEL` – model name to send in the request body
- `ENABLE_LLM_EXPLAIN` – enables/disables AI interpretation globally (`"true"` or `"false"`)
- `LLM_TIMEOUT` – request timeout in seconds (default: `45`)

If these values are missing or invalid, the app will fall back to its built-in rule-based interpretation instead of failing completely.

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
```

You can also set the same values as environment variables instead of using `secrets.toml`.

**Important:** do not commit `.streamlit/secrets.toml` to Git.

#### Streamlit Community Cloud

For the deployed app, update the same values in:

**App Settings → Secrets**

Paste the same block into the Secrets editor:

```toml
LLM_API_KEY = "your-new-api-key"
LLM_BASE_URL = "https://api.openai.com/v1/chat/completions"
LLM_MODEL = "gpt-4.1"
ENABLE_LLM_EXPLAIN = "true"
LLM_TIMEOUT = "45"
```

After saving, restart or refresh the app and run a simulation to confirm the new key is working.

#### Rotating or changing the API key

To replace the current LLM key:

1. Create a new API key with your preferred provider/account.
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




## Local development only — External venv (Windows)

**For local development on your machine only.** Use this if creating `.venv` inside the project fails (e.g. path length, antivirus, or file-copy errors). The virtual environment lives outside the repo at `C:\\venvs\\ip-app`.

1. Create the parent folder once if it does not exist:

```powershell
New-Item -ItemType Directory -Force -Path C:\\venvs
```

2. Create the venv (once), activate it, go to the project, install, and run:

```powershell
py -m venv C:\\venvs\\ip-app
C:\\venvs\\ip-app\\Scripts\\Activate.ps1
cd "D:\\Niharika\\Purdue BAIM Courses\\Spring 26\\Mod 4\\IP"
pip install -r requirements.txt
streamlit run app.py
```

Adjust the `cd` path if your project folder is elsewhere.
3. **Each new terminal session:** activate the venv, then `cd` to the project before `pip` or `streamlit`:

```powershell
C:\\venvs\\ip-app\\Scripts\\Activate.ps1
cd "D:\\Niharika\\Purdue BAIM Courses\\Spring 26\\Mod 4\\IP"
streamlit run app.py
```


## Usage

1. **Select products** from the sidebar (Archetype | Maturity combinations).
2. **Choose a production strategy** -- custom per-year multipliers, just-in-time, conservative, or aggressive.
3. **Configure simulation settings** -- number of iterations and random seed.
4. **Set analysis parameters** -- year mode, remaining inventory threshold, and production constraints.
5. Click **Run Simulation** to execute the Monte Carlo engine.
6. View the mean lifecycle table, probability summary, key metric cards, and remaining inventory chart in the main area.

Yes, this is a solid start, but for GDM and for a public repo you want:

- Clear business/problem context.
- A quick overview of methodology.
- A short “repo structure” section.
- Slightly cleaned usage language that matches the current app.[^1]

Below is a revised version you can more or less paste in and tweak to match any small app changes.

***

# GDM Decision Simulation Tool

A Streamlit dashboard for running multi‑year Monte Carlo simulations of seed production, sales, and inventory. The tool models yield and conversion uncertainty, lifecycle‑driven demand, and inventory carryover to quantify stockout and excess‑inventory risk for different production strategies.[^1]

## Project objective

This application was developed as part of the Spring 2026 Industry Practicum with GDM/AgReliant. It is designed to help planners:

- Evaluate production strategies over a 10‑year product lifecycle.
- Quantify stockout risk and carryover inventory under historical volatility.
- Compare outcomes across archetype–maturity segments and strategies (e.g., 1.0×–2.0× multipliers).[^2][^1]


## Repository structure

- `app.py` – Main Streamlit dashboard application (UI and simulation engine).
- `.streamlit/` – Streamlit configuration (layout, theme, etc.).
- `validation.py` – Validation / QA utilities for checking simulation outputs and assumptions.
- `requirements.txt` – Python dependencies.
- `GDM_Mar17.ipynb` – Early exploratory notebook / prototype work.
- `GDM_final.ipynb` – Final analysis notebook used to design and test the production/inventory logic.
- `GDM_final (2).ipynb` – Alternative export of the final notebook (can be cleaned up or renamed).
- `IP Carryover Model 1.xlsx` – Excel prototype of the carryover logic used during model design.
- `README.md` – Project documentation (this file).[^1]

If you rename or remove legacy notebooks, update this section accordingly.

## Data inputs

By default the app expects the following files in the project root:

- `data.xlsx`
    - Production yields
    - Conversion rates
    - Product parameters
    - Median / average sales parameters
- `salesVariability.xlsx`
    - Lognormal parameters for first‑year sales
    - Lognormal parameters for year‑over‑year growth/decline rates

These files are derived from GDM historical data (2015–2025) and define the distributions the simulation draws from.[^2]

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


### 3. Run the app

```bash
streamlit run app.py
```

The app will open in your browser at `http://localhost:8501`.

## Optional: external venv for local Windows development

If creating `.venv` inside the project fails (e.g., path length or antivirus restrictions), you can create an external environment such as `C:\venvs\ip-app`:

```powershell
New-Item -ItemType Directory -Force -Path C:\venvs

py -m venv C:\venvs\ip-app
C:\venvs\ip-app\Scripts\Activate.ps1
cd "D:\Niharika\Purdue BAIM Courses\Spring 26\Mod 4\IP"  # adjust path
pip install -r requirements.txt
streamlit run app.py
```

For each new terminal session:

```powershell
C:\venvs\ip-app\Scripts\Activate.ps1
cd "D:\Niharika\Purdue BAIM Courses\Spring 26\Mod 4\IP"
streamlit run app.py
```


## How the tool is used

1. **Select products** from the sidebar (archetype–maturity combinations matching GDM’s taxonomy).
2. **Choose a production strategy** (e.g., 1.0× baseline, conservative, aggressive 2.0×).
3. **Configure simulation settings** – number of Monte Carlo iterations and random seed for reproducibility.
4. **Set analysis parameters** – analysis year (e.g., Year 6), inventory threshold, and any production constraints.
5. Click **Run simulation** to execute the 10‑year lifecycle engine.
6. Review:
    - Year‑by‑year inventory table (carry‑in, production, sales, remaining).
    - Summary metrics (mean/median total sales, mean/median remaining inventory).
    - Risk metrics (`P(remaining > 0)` and `P(depleted)` at the selected year and at Year 10).[^2]

## Methodology (summary)

- Yield and conversion variability are estimated from historical production data and modeled as random draws around planned values.
- Lifecycle demand is based on archetype–maturity sales curves (first‑year volume plus YOY growth/decline factors).
- Inventory evolves year‑by‑year with production, demand, 2% production quality loss, and 10% annual carryover degradation.
- Monte Carlo simulation is used to generate many possible futures; the dashboard summarizes these paths into expected values and risk probabilities to support planning decisions.[^2]
