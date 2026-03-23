# Monte Carlo Production & Inventory Planner

A Streamlit dashboard for running 10-year Monte Carlo simulations of seed production and inventory across multiple products. Uses lognormal sales variability, yield & conversion uncertainty, and portfolio-level aggregation.

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
.venv\Scripts\Activate.ps1
```

**Windows (Command Prompt):**

```cmd
python -m venv .venv
.venv\Scripts\activate.bat
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

### 3. Run the app

```
streamlit run app.py
```

The app will open in your browser at [http://localhost:8501](http://localhost:8501).

## Local development only — External venv (Windows)

**For local development on your machine only.** Use this if creating `.venv` inside the project fails (e.g. path length, antivirus, or file-copy errors). The virtual environment lives outside the repo at `C:\venvs\ip-app`.

1. Create the parent folder once if it does not exist:

   ```powershell
   New-Item -ItemType Directory -Force -Path C:\venvs
   ```

2. Create the venv (once), activate it, go to the project, install, and run:

   ```powershell
   py -m venv C:\venvs\ip-app
   C:\venvs\ip-app\Scripts\Activate.ps1
   cd "D:\Niharika\Purdue BAIM Courses\Spring 26\Mod 4\IP"
   pip install -r requirements.txt
   streamlit run app.py
   ```

   Adjust the `cd` path if your project folder is elsewhere.

3. **Each new terminal session:** activate the venv, then `cd` to the project before `pip` or `streamlit`:

   ```powershell
   C:\venvs\ip-app\Scripts\Activate.ps1
   cd "D:\Niharika\Purdue BAIM Courses\Spring 26\Mod 4\IP"
   streamlit run app.py
   ```

## Usage

1. **Select products** from the sidebar (Archetype | Maturity combinations).
2. **Choose a production strategy** -- custom per-year multipliers, just-in-time, conservative, or aggressive.
3. **Configure simulation settings** -- number of iterations and random seed.
4. **Set analysis parameters** -- year mode, remaining inventory threshold, and production constraints.
5. Click **Run Simulation** to execute the Monte Carlo engine.
6. View the mean lifecycle table, probability summary, key metric cards, and remaining inventory chart in the main area.
