# DaveAI — Part 1: Foundations, GitHub & the Data Pipeline

> **DaveAI** = **D**eterministic **A**vicollect **V**irtual **E**xpert.
> A local, privacy‑preserving chatbot that answers plain‑English questions about historic flight data — *including* real‑math questions like "total flights in Q1 2023" or "average passenger count in March 2025" — by translating your question into a real SQL query and letting a real database do the arithmetic.

This is **Part 1 of 3**. When you finish it you will have: a working local AI runtime on your Intel Arc GPU, a clean GitHub repo, and a data pipeline that loads your two Excel files into a real database you can query. Parts 2 and 3 (the actual question‑answering brain, MLflow, and voice/UI) build directly on top of what you set up here — the folder layout below is designed so they slot in without any restructuring.

A note on the audience: this guide assumes you can write code but have **never** built an AI pipeline, connected an LLM to a database, or used Ollama, DuckDB, or MLflow. Every tool is explained in plain language before you install it, and every block of code is explained line‑by‑idea, not just dumped.

---

## 0. The big picture — how DaveAI actually works

Before any setup, understand the one idea the whole project rests on. There are two ways an AI could answer *"how many flights were there in Q1 2023?"*:

- **The naïve way (bad):** hand the raw data to a large language model (LLM) and ask it to count. LLMs are pattern generators, not calculators. They *guess* plausible‑looking numbers. This is called **hallucination**, and for anything involving money or counts it is unacceptable.
- **The DaveAI way (good):** the LLM never touches the numbers. Instead, it **translates your English question into SQL** (the language databases speak). That SQL runs against a real database, and the *database* computes the answer. The LLM's only job is translation — turning "flights in Q1 2023" into `SELECT COUNT(*) ... WHERE flight_date BETWEEN ...`.

This is the **text‑to‑SQL** architecture, and it is exactly why the project is named *Deterministic*: the same question always produces the same SQL and therefore the same, verifiable number. The answer is *grounded by SQL execution*, not conjured by free‑form text generation.

The flow, end to end:

```
Your question (English)
        │
        ▼
   ┌─────────┐   writes SQL    ┌──────────┐   runs SQL   ┌─────────┐
   │  LLM    │ ──────────────▶ │  DuckDB  │ ───────────▶ │ Result  │
   │ (Qwen)  │                 │ database │              │ (numbers)│
   └─────────┘                 └──────────┘              └─────────┘
        ▲                                                     │
        │                 explains result in English          │
        └─────────────────────────────────────────────────────┘
```

**Part 1 builds the right‑hand side of that picture** — the database and the clean data inside it — plus the local LLM runtime that Part 2 will wire in as the "writes SQL" step.

---

## 1. The tools, in plain language

You'll hear these names constantly, so here's what each one *is* and *why it's in the stack* before you install anything.

**Python virtual environment (`venv`)** — A private, isolated copy of Python for this project only. Instead of installing libraries globally (where versions from different projects collide), you install them into a folder that belongs to DaveAI alone. Delete the folder and it's like the install never happened. This is standard hygiene for any serious Python project.

**Ollama** — Think of Ollama as "Docker for language models." It's a small program that downloads open‑weight LLMs, runs them locally on your machine, and exposes them through a simple local web address (`http://localhost:11434`). No cloud, no API keys, no data leaving your computer — which is the whole point of a privacy‑preserving assistant. You talk to it either from the command line (`ollama run ...`) or from Python.

**Qwen2.5‑7B‑Instruct** — The actual "brain": an open‑weight LLM made by Alibaba Cloud, released under the permissive **Apache 2.0 license** (free to use, including commercially). "7B" means 7 billion parameters — big enough to write good SQL, small enough (~4.7 GB) to run on your hardware. "Instruct" means it's tuned to follow instructions and chat, rather than just autocomplete text. Qwen2.5 is specifically strong at structured data and generating structured output, which is ideal for a text‑to‑SQL job.

**DuckDB** — An **embedded analytical database**. "Embedded" means there's no server to install or manage — the entire database is a single file on disk (`daveai.duckdb`), and you talk to it directly from Python. "Analytical" means it's built for exactly the kind of query DaveAI generates: `SUM`, `AVG`, `COUNT`, grouping by month/airline/route across many rows. It's astonishingly fast at this and needs essentially zero configuration. It's to analytics what SQLite is to app data.

**pandas + openpyxl** — **pandas** is Python's standard library for working with tables of data (its core object, the *DataFrame*, is basically a spreadsheet in code). **openpyxl** is the engine pandas uses under the hood to physically read `.xlsx` files. You need both: pandas to clean and reshape the data, openpyxl so pandas can open Excel at all.

**Vulkan (and why it matters for *your* machine)** — This is the one hardware‑specific piece, so read carefully. To run an LLM fast, you offload the math onto a GPU. NVIDIA GPUs do this through **CUDA** — but **your Intel Arc Pro 140T does not support CUDA**. Intel GPUs instead use **Vulkan**, a cross‑vendor graphics/compute API. Recent versions of Ollama include a Vulkan backend, which is what lets your Intel iGPU accelerate inference. Two important, current facts:

- On **Windows**, Vulkan support ships *inside the Intel Arc graphics driver*. You do **not** need to install a separate "Vulkan runtime" just to *run* models — keeping your GPU driver up to date is enough. (A separate Vulkan SDK is only needed if you ever compile Ollama from source, which you won't here.)
- You may find older tutorials that tell you to install **IPEX‑LLM**, an Intel‑specific fork of Ollama. **Ignore those.** Intel archived the IPEX‑LLM project in January 2026. The native Vulkan path in mainstream Ollama is now the correct, supported way to use an Intel Arc GPU.

**plotly and matplotlib** — Two Python charting libraries. You won't draw anything in Part 1, but you'll install them now (they're in `requirements.txt`) because Parts 2 and 3 use them to visualize results, and it's cleaner to pin every dependency once, up front.

---

## 2. Install everything (Windows, step by step)

### 2.1 Install Python and create the virtual environment

1. If you don't already have Python, install **Python 3.12** from [python.org/downloads](https://www.python.org/downloads/). On the first screen of the installer, **tick "Add python.exe to PATH"** before clicking Install — this saves a lot of pain later.

2. Open **PowerShell** and confirm it worked:
   ```powershell
   python --version
   ```
   You should see something like `Python 3.12.x`. If PowerShell says it doesn't recognize `python`, reopen PowerShell (PATH changes need a fresh window) or reinstall with the PATH box ticked.

3. Choose a folder for the project and create it. We'll clone the GitHub repo into this same place in §3, so pick somewhere sensible:
   ```powershell
   cd $HOME\Documents
   mkdir dave-ai
   cd dave-ai
   ```

4. Create the virtual environment (the isolated Python) and activate it:
   ```powershell
   python -m venv .venv
   .\.venv\Scripts\Activate.ps1
   ```
   `python -m venv .venv` builds the environment in a hidden folder called `.venv`. The second line *activates* it. Your prompt should now start with `(.venv)` — that's how you know you're inside it. **You must activate the venv every time you open a new terminal to work on this project.**

   > If activation is blocked by an error about "running scripts is disabled on this system," run PowerShell once with:
   > ```powershell
   > Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
   > ```
   > then try activating again. This is a standard, safe one‑time setting.

### 2.2 Install the Python libraries

We'll create the full `requirements.txt` file properly in §7, but you can install the pinned set now. Create a file named `requirements.txt` in the project folder with the contents shown in §7, then:

```powershell
python -m pip install --upgrade pip
pip install -r requirements.txt
```

The first line updates pip (Python's package installer) itself. The second reads `requirements.txt` and installs every library at the exact pinned version. Pinning versions means the project behaves identically on any machine and won't silently break when a library releases an update.

### 2.3 Install the Ollama application

**Important distinction to avoid confusion later:** there are *two* different things called "ollama":
- the **Ollama application** — the program that actually runs models, installed from the website below; and
- the **`ollama` Python package** — a small client library (already installed via `requirements.txt`) that lets your Python code *talk* to the application.

They have separate version numbers. Part 2 uses the Python package; right now we install the application.

1. Download the Windows installer from **[ollama.com/download/windows](https://ollama.com/download/windows)** and run it. It installs as a normal Windows program (no WSL, no Docker needed) and starts a small background service plus a llama icon in your system tray.

2. Confirm it's installed by opening a **new** PowerShell window and running:
   ```powershell
   ollama --version
   ```
   You should get a version string (mid‑2026 releases are `0.x` in the 0.30+ range). The exact number isn't critical, but **newer is better** here because Vulkan/Intel support has improved steadily — if yours is old, reinstall the latest.

### 2.4 Update your Intel Arc graphics driver

Because Vulkan lives in the driver, an up‑to‑date driver is what makes GPU acceleration work at all. Install the latest **Intel Arc & Iris Xe graphics driver** from [intel.com](https://www.intel.com/content/www/us/en/download-center/home.html) (or via the Intel Graphics Software app), then **reboot**. Your Arc Pro 140T with ~18 GB of shared memory is fully Vulkan‑capable, so once the driver is current, Ollama can see it.

### 2.5 Download the model

Pull Qwen2.5‑7B‑Instruct. "Pull" means download‑and‑register, exactly like `docker pull`:

```powershell
ollama pull qwen2.5:7b-instruct
```

This fetches ~4.7 GB (the Q4_K_M quantization — a compressed format that keeps quality high while shrinking the file to fit comfortably in your memory). It runs once; afterwards the model lives on your disk. Use this exact tag — `qwen2.5:7b-instruct` — everywhere, so you always get the instruction‑tuned 7B model.

### 2.6 First run — and *proving* the GPU (not the CPU) is doing the work

Start a quick chat to confirm the model works at all:

```powershell
ollama run qwen2.5:7b-instruct "In one sentence, what is SQL?"
```

It should reply in a few seconds. Now the part your prompt specifically asked about — **verifying it's using the Arc GPU via Vulkan and not silently falling back to CPU.** Use all three checks:

**Check A — `ollama ps` (the quickest signal).** While a model is loaded (run the command above first, or start `ollama run qwen2.5:7b-instruct` and leave it open in another window), run:
```powershell
ollama ps
```
Look at the **PROCESSOR** column:
- `100% GPU` → fully accelerated. This is what you want.
- `100% CPU` → it's running on the CPU only (the thing we're trying to avoid).
- `48%/52% CPU/GPU` → partially offloaded. For a 4.7 GB model on 18 GB of memory this shouldn't happen; if it does, something's misconfigured.

**Check B — Task Manager.** Open Task Manager → **Performance** → **GPU** (select the Intel Arc GPU if several are listed). Send the model a prompt and watch: GPU utilization should jump to 60–100% during the reply. If GPU stays flat at 0% while a CPU core pegs at 100%, it's on the CPU.

**Check C — the debug log (most detailed).** In PowerShell:
```powershell
$env:OLLAMA_DEBUG="1"; ollama run qwen2.5:7b-instruct "hello"
```
Scan the startup log for a line naming the backend and device. Seeing **Vulkan** and your Intel GPU mentioned confirms the GPU path; seeing only CPU/threads and no GPU backend confirms a CPU fallback.

### 2.7 If it fell back to CPU — troubleshooting

Work through these in order:

1. **Update the Intel driver and reboot** (§2.4). A stale or broken driver is the single most common cause; a driver update that breaks Vulkan is also possible, in which case rolling back one version can fix it.
2. **Update Ollama** to the latest release. Vulkan support and Intel iGPU detection have improved release‑over‑release.
3. **Force Vulkan on explicitly.** Some builds treat it as opt‑in. Set the environment variable and restart Ollama:
   ```powershell
   setx OLLAMA_VULKAN 1
   ```
   Then fully quit Ollama from the tray and reopen it (environment changes need a restart), and re‑check with `ollama ps`.
4. **Keep the model resident** so you're not paying the load cost repeatedly while testing:
   ```powershell
   setx OLLAMA_KEEP_ALIVE 30m
   ```

**A realistic expectation:** an integrated GPU shares system memory (LPDDR5X) rather than having dedicated high‑bandwidth VRAM, so it's *memory‑bandwidth‑bound*. GPU acceleration here is a real and worthwhile speedup over CPU, but a 7B model on an iGPU is "comfortably usable," not "blazing." That's fine — DaveAI generates short SQL queries, not essays.

> **Fallback option (optional):** if you ever can't get Vulkan working through Ollama, **LM Studio** on Windows has very stable Vulkan support for Intel Arc and can serve an Ollama‑compatible local API. We stay on Ollama for this guide as planned, but it's good to know the escape hatch exists.

---

## 3. GitHub from scratch

If you already use Git and GitHub daily, skim this; it's written for someone doing it deliberately for the first time on this project.

### 3.1 Create the repository on GitHub

1. Sign in at [github.com](https://github.com). Click **New** (green button) to create a repository.
2. **Repository name:** `dave-ai`. **Description:** *"DaveAI — Deterministic Avicollect Virtual Expert: a local text‑to‑SQL chatbot for historic flight data."*
3. Set it **Private** (this is flight/financial data tooling). Do **not** tick "Add a README" or "Add .gitignore" — you'll create better ones locally in a moment.
4. Click **Create repository** and leave the page open; you'll need the URL it shows.

### 3.2 Initialize Git locally and connect it

Make sure you're in the project folder (`dave-ai`) with the venv active, then:

```powershell
git init
git branch -M main
git remote add origin https://github.com/<your-username>/dave-ai.git
```

`git init` turns the folder into a Git repository. `git branch -M main` names the primary branch `main`. `git remote add origin ...` links your local repo to the GitHub one (replace `<your-username>`).

### 3.3 Write a proper `.gitignore`

A `.gitignore` tells Git which files to **never** upload. This matters enormously here: your raw flight data and your built database must **stay off GitHub** (privacy, and they're large/regenerable). Create a file named exactly `.gitignore` with this content:

```gitignore
# --- Python ---
.venv/
__pycache__/
*.py[cod]
*.egg-info/
.ipynb_checkpoints/

# --- Secrets / config ---
.env

# --- DaveAI data (NEVER commit raw data or the built DB) ---
data/raw/*
!data/raw/.gitkeep
data/*.duckdb
*.duckdb

# --- Part 2/3 outputs (safe to ignore now) ---
mlruns/
models/
*.log
```

What each block does: the Python section skips your virtual environment, compiled caches, and notebook checkpoints. The secrets line skips a future `.env` file. The **data section is the critical one** — `data/raw/*` ignores every raw Excel file, while `!data/raw/.gitkeep` makes an exception so the *empty folder* is still tracked (Git otherwise won't track empty folders); `*.duckdb` ignores the built database anywhere it appears. The last block pre‑ignores things Parts 2 and 3 will create (MLflow runs, model artifacts, logs) so you never accidentally commit them.

### 3.4 Write the starter `README.md`

Create `README.md`:

```markdown
# DaveAI

**DaveAI** stands for **D**eterministic **A**vicollect **V**irtual **E**xpert — a local,
privacy‑preserving chatbot that answers natural‑language questions about historic flight
data. It uses a **text‑to‑SQL** architecture: your question is translated into a real SQL
query that runs against a local DuckDB database, so all arithmetic is computed by the
database, not guessed by the language model.

The name is deliberate. "Deterministic" reflects the architecture: answers are *grounded
by SQL execution*, not produced by free‑form text generation — the LLM never invents a
number, the database always computes it. Nothing leaves your machine: no cloud APIs, no
external data sharing.

## Stack
- **LLM runtime:** Ollama, running Qwen2.5‑7B‑Instruct (Apache 2.0, open‑weight)
- **Database:** DuckDB (embedded analytical database)
- **Language:** Python
- **Excel:** pandas + openpyxl
- **GPU:** Intel Arc (Vulkan backend — no CUDA)

## Status
- [x] Part 1 — Environment, GitHub, and data pipeline
- [ ] Part 2 — Text‑to‑SQL core + MLflow experiment tracking
- [ ] Part 3 — Voice I/O (English, Yoruba, Igbo, Hausa) + UI + deployment

## Quick start
See the Part 1 guide. In short: create the venv, `pip install -r requirements.txt`,
put your Excel files in `data/raw/`, then `python src/db_loader.py`.
```

### 3.5 First commit and push

```powershell
git add .
git commit -m "Part 1: project scaffold, data pipeline, and setup guide"
git push -u origin main
```

`git add .` stages every file *except* the ones `.gitignore` excludes (verify raw data isn't staged with `git status` — you should **not** see your `.xlsx` files listed). `git commit` saves a snapshot with a message. `git push -u origin main` uploads it to GitHub. GitHub may prompt you to sign in the first time.

---

## 4. Project structure

Create these folders and empty files now. This exact layout is chosen so that Part 2 (`pipeline/`, `mlruns/`) and Part 3 (`voice/`, `app/`) drop in with **no restructuring**.

```
dave-ai/
├── .venv/                     # virtual environment (git‑ignored)
├── data/
│   ├── raw/                   # YOUR two Excel files go here (git‑ignored)
│   │   └── .gitkeep           # keeps the empty folder tracked by Git
│   └── daveai.duckdb          # the built database (git‑ignored, created by code)
├── src/
│   ├── __init__.py            # marks src as a Python package
│   ├── config.py              # one place for paths + the model name
│   ├── excel_loader.py        # reads & cleans the Excel files
│   ├── db_loader.py           # loads clean data into DuckDB
│   ├── schema.py              # plain‑English column docs for the LLM
│   ├── pipeline/              # (Part 2) text‑to‑SQL core + MLflow
│   │   └── __init__.py
│   ├── voice/                 # (Part 3) speech‑to‑text / text‑to‑speech
│   │   └── __init__.py
│   └── app/                   # (Part 3) the user interface
│       └── __init__.py
├── scripts/
│   └── verify_load.py         # a runnable proof that data loaded
├── notebooks/                 # exploratory Jupyter notebooks
├── mlruns/                    # (Part 2) MLflow will write here (git‑ignored)
├── .gitignore
├── requirements.txt
└── README.md
```

**What belongs where:** `data/raw/` holds your untouched source Excel files (and nothing else); the database is written next to it. `src/` is all your importable Python code — the loaders and schema live at its top level, and the empty `pipeline/`, `voice/`, `app/` sub‑packages are placeholders reserved for later parts. `scripts/` holds runnable one‑off tools like the verifier. `notebooks/` is for experimentation. `mlruns/` and `models/` are created later by Part 2. An `__init__.py` file (even empty) is what makes a folder a *package* Python can import from — that's why each code folder has one.

Create the empty scaffolding quickly from PowerShell:

```powershell
mkdir data\raw, src\pipeline, src\voice, src\app, scripts, notebooks, mlruns
New-Item data\raw\.gitkeep, src\__init__.py, src\pipeline\__init__.py, src\voice\__init__.py, src\app\__init__.py -ItemType File
```

---

## 5. The data pipeline code

Now the heart of Part 1. **First, copy your two files into `data/raw/`** with these exact names:
- `data/raw/manifest_report.xlsx` (one row per flight)
- `data/raw/passenger_data.xlsx` (one row per passenger)

All the code below has been tested end‑to‑end against files matching your described columns and example rows (comma‑formatted currency like `"315,794.90"`, the two different date styles, and the empty itinerary‑leg columns).

### 5.1 `src/config.py` — one home for paths and the model name

Rather than scatter file paths through every file, keep them in one place. Parts 2 and 3 will import from here too.

```python
"""config.py -- one place for paths and the model name.
Parts 2 (MLflow) and 3 (voice/UI) will import from here too."""
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
DATA_DIR = PROJECT_ROOT / "data"
RAW_DIR = DATA_DIR / "raw"
DB_PATH = DATA_DIR / "daveai.duckdb"

MANIFEST_FILE = RAW_DIR / "manifest_report.xlsx"
PASSENGER_FILE = RAW_DIR / "passenger_data.xlsx"

OLLAMA_MODEL = "qwen2.5:7b-instruct"
```

**Why it's written this way:** `Path(__file__).resolve().parent.parent` finds the project root *relative to this file*, so the paths are correct no matter which folder you run your scripts from. Every other module asks `config` for paths instead of hard‑coding them — change a path once here and the whole project follows.

### 5.2 `src/excel_loader.py` — read and clean the Excel files

This is where messy spreadsheet reality gets turned into clean, correctly‑typed data. The golden rule: **the database can only do correct math if numbers are real numbers and dates are real dates.** If `"315,794.90"` stays as text, no `SUM` will ever work on it.

```python
"""excel_loader.py -- DaveAI (Deterministic Avicollect Virtual Expert)
Reads the two raw Excel workbooks into clean pandas DataFrames."""
import pandas as pd
import config

# passenger_data dates look like "09/02/2023". Slash dates are AMBIGUOUS:
# is that 9 Feb or 2 Sep? If any day value in the column exceeds 12, the
# format must be day-first (DD/MM). Set this to match YOUR data (see note below).
PASSENGER_DATE_DAYFIRST = True


def clean_currency(series: pd.Series) -> pd.Series:
    """Turn currency-formatted text like '315,794.90' into the float 315794.90.
    Blank / dash cells become NaN (missing) -- NOT 0, because 0 is a real amount."""
    cleaned = (
        series.astype("string")
        .str.replace(",", "", regex=False)   # remove thousands separators
        .str.replace("N", "", regex=False)   # strip stray naira letters if present
        .str.replace("=", "", regex=False)   # strip '=N=' style markers
        .str.strip()
        .replace({"": pd.NA, "-": pd.NA, "nan": pd.NA})
    )
    return pd.to_numeric(cleaned, errors="coerce")   # anything unparseable -> NaN


def parse_dates(series: pd.Series, dayfirst: bool = False) -> pd.Series:
    """Parse a text date column into real datetimes. Unparseable cells -> NaT."""
    return pd.to_datetime(series, dayfirst=dayfirst, errors="coerce")


def load_manifest() -> pd.DataFrame:
    # dtype=str: read EVERY cell as text first, so pandas can't guess wrong
    # (e.g. reading a flight number as a number and dropping a leading zero).
    df = pd.read_excel(config.MANIFEST_FILE, engine="openpyxl", dtype=str)
    df.columns = [c.strip() for c in df.columns]   # trim stray spaces in headers

    df["Pax Count"] = pd.to_numeric(df["Pax Count"], errors="coerce").astype("Int64")
    df["TSC Amount(N)"] = clean_currency(df["TSC Amount(N)"])
    df["TSC Amount($)"] = clean_currency(df["TSC Amount($)"])
    df["Flight Date"] = parse_dates(df["Flight Date"])       # "30 Jun, 2024"
    df["Date Submitted"] = parse_dates(df["Date Submitted"])

    # Rename to clean, SQL-friendly column names (no spaces, no symbols).
    return df.rename(columns={
        "S/N": "sn", "Airline": "airline", "Flight Number": "flight_number",
        "Route": "route", "Status": "status", "Pax Count": "pax_count",
        "Flight Date": "flight_date", "Date Submitted": "date_submitted",
        "TSC Amount(N)": "tsc_amount_ngn", "TSC Amount($)": "tsc_amount_usd",
    })


def load_passengers() -> pd.DataFrame:
    df = pd.read_excel(config.PASSENGER_FILE, engine="openpyxl", dtype=str)
    df.columns = [str(c).strip() for c in df.columns]

    # The 6 itinerary legs arrive as columns literally named "1".."6".
    # Rename them to leg_1..leg_6 -- purely numeric names confuse SQL and the LLM.
    df = df.rename(columns={str(i): f"leg_{i}" for i in range(1, 7)})
    for i in range(1, 7):
        col = f"leg_{i}"
        # Empty leg = "this leg wasn't used", which is a true NULL, not bad data.
        df[col] = df[col].astype("string").str.strip().replace({"": pd.NA})

    df["DATE OF OPERATION"] = parse_dates(
        df["DATE OF OPERATION"], dayfirst=PASSENGER_DATE_DAYFIRST
    )
    for c in ["FARE USD ($)", "FARE NGN (=N=)", "SALES LEVY ($)", "SALES LEVY (=N=)"]:
        df[c] = clean_currency(df[c])
    df["FREE TICKET"] = pd.to_numeric(df["FREE TICKET"], errors="coerce").astype("Int64")

    return df.rename(columns={
        "S/N": "sn", "Airline": "airline", "DATE OF OPERATION": "date_of_operation",
        "FLIGHT NUMBER": "flight_number", "PAX NAME": "pax_name",
        "PAX TYPE": "pax_type", "TICKET NO/ENDORSED TICKET": "ticket_no",
        "FARE TYPE": "fare_type", "Currency of Sale": "currency_of_sale",
        "FARE USD ($)": "fare_usd", "FARE NGN (=N=)": "fare_ngn",
        "SALES LEVY ($)": "sales_levy_usd", "SALES LEVY (=N=)": "sales_levy_ngn",
        "FREE TICKET": "free_ticket",
    })
```

**The cleaning decisions, explained:**

- **`dtype=str` on read.** We deliberately read *everything* as text first. If you let pandas auto‑detect types, it can silently mangle values — turning a flight number like `0912` into `912`, or misreading a currency string. Reading as text, then converting *on purpose*, keeps you in control.
- **`clean_currency`.** Your amounts arrive as strings with thousands separators (`"717,955.00"`). Removing the commas (and any stray `N`/`=` naira markers) and calling `pd.to_numeric` yields a real number the database can sum. Crucially, blanks become `NaN` (missing) rather than `0` — a missing fare and a genuinely‑zero fare are different facts, and conflating them would corrupt any average.
- **Two date formats, handled separately.** The manifest uses `"30 Jun, 2024"`, which pandas parses unambiguously. The passenger file uses `"09/02/2023"`, which is **ambiguous** (9 Feb or 2 Sep?). The `PASSENGER_DATE_DAYFIRST` switch controls it. **How to check which is right for your data:** scan the `DATE OF OPERATION` column for any value where the first number is greater than 12 (e.g. `13/07/2024`). If you find one, the format is day‑first → keep `True`. If the *second* number ever exceeds 12, it's month‑first → set `False`. Getting this wrong flips days and months silently, so verify it once.
- **The leg columns.** Columns named `1`–`6` become `leg_1`–`leg_6`. Numeric column names are poison for text‑to‑SQL: the LLM can't tell `"1"` (a column) from `1` (the number), and SQL needs awkward quoting. Renaming also lets Part 2's LLM understand that `leg_1` is the departure airport. Empty legs become true `NULL`s.

### 5.3 `src/db_loader.py` — load the clean data into DuckDB

```python
"""db_loader.py -- loads the cleaned DataFrames into a local DuckDB database.

DuckDB is an embedded analytical database: it lives in one .duckdb file on your
machine, needs no server, and is fast at the SUM/AVG/COUNT-by-group queries
DaveAI will generate. Nothing leaves your computer."""
import duckdb
import config
from excel_loader import load_manifest, load_passengers


def load_dataframe(con, df, table_name: str) -> None:
    """Register a pandas DataFrame with DuckDB, then copy it into a real table."""
    view = f"_tmp_{table_name}"
    con.register(view, df)          # expose the DataFrame to SQL under a temp name
    con.execute(f"CREATE OR REPLACE TABLE {table_name} AS SELECT * FROM {view}")
    con.unregister(view)            # tidy up the temporary registration


def build_database():
    con = duckdb.connect(str(config.DB_PATH))     # creates the .duckdb file if missing
    load_dataframe(con, load_manifest(), "manifest_reports")
    load_dataframe(con, load_passengers(), "passenger_records")
    con.close()
    return config.DB_PATH


if __name__ == "__main__":
    print("Database written to:", build_database())
```

**What's happening:** `duckdb.connect(...)` opens (or creates) the database file. `con.register(...)` momentarily lets DuckDB see a pandas DataFrame as if it were a table; `CREATE OR REPLACE TABLE ... AS SELECT * FROM ...` then copies that data into a *permanent* table stored in the file. We do this for both DataFrames, producing two tables — `manifest_reports` and `passenger_records`. `CREATE OR REPLACE` means re‑running the loader safely rebuilds the tables from scratch rather than erroring or duplicating rows. The `if __name__ == "__main__":` block means the file *builds the database when you run it directly* but stays quiet when imported by other code.

Run it (from the project root, venv active):

```powershell
python src\db_loader.py
```

Expected output:

```
Database written to: ...\dave-ai\data\daveai.duckdb
```

### 5.4 `src/schema.py` — teach the LLM what your columns *mean*

This file looks humble but is one of the most important in the whole project. In Part 2, the LLM writes SQL — but it can only write *correct* SQL if it knows what each column means. It has no way to guess that `tsc_amount_ngn` is a charge in naira, or that `pax_type` uses letter codes, or that `leg_1` is the departure airport. `schema.py` is the plain‑English "data dictionary" you'll feed into the LLM's prompt so its SQL is grounded in your actual business meaning. **Better descriptions here directly produce more accurate answers.**

```python
"""schema.py -- a plain-English description of every table and column.

This text is injected into the LLM's prompt in Part 2 so it can write correct
SQL. Accuracy of these descriptions DIRECTLY drives the accuracy of DaveAI's
answers -- if a description is vague or wrong, the generated SQL will be too.
Fill in the [CONFIRM] notes with your exact business meaning."""

SCHEMA_TEXT = """
You are querying a local DuckDB database about historic flight data.
There are two tables.

TABLE: manifest_reports  (one row = one flight)
  sn              INTEGER  - serial number from the source file (not analytical)
  airline         TEXT     - name or code of the operating airline
  flight_number   TEXT     - flight designator, e.g. 'Q9303'
  route           TEXT     - origin-destination airport codes joined by '-',
                             e.g. 'ABV-LOS' means Abuja to Lagos
  status          TEXT     - manifest submission status, e.g. 'Submitted',
                             'Not Submitted'
  pax_count       INTEGER  - number of passengers on the flight
  flight_date     DATE     - date the flight operated
  date_submitted  DATE     - date the manifest was submitted
  tsc_amount_ngn  DOUBLE   - TSC charge for the flight in Nigerian Naira. [CONFIRM
                             the exact meaning of 'TSC' for your business]
  tsc_amount_usd  DOUBLE   - TSC charge for the flight in US Dollars

TABLE: passenger_records  (one row = one passenger)
  sn                TEXT    - serial number from the source file (not analytical)
  airline           TEXT    - airline name or code
  date_of_operation DATE    - date this passenger's flight operated
  flight_number     TEXT    - flight designator
  pax_name          TEXT    - passenger full name
  pax_type          TEXT    - passenger type code, e.g. 'A'=Adult, 'C'=Child,
                              'I'=Infant. [CONFIRM your exact codes]
  ticket_no         TEXT    - ticket number / endorsed ticket number
  leg_1             TEXT    - itinerary leg 1: the DEPARTURE airport code
  leg_2             TEXT    - itinerary leg 2 (destination of a 2-leg trip, or a
                              connecting stop if more legs)
  leg_3             TEXT    - itinerary leg 3: final outbound destination
  leg_4             TEXT    - itinerary leg 4: return-leg airport (NULL if unused)
  leg_5             TEXT    - itinerary leg 5: return-leg airport (NULL if unused)
  leg_6             TEXT    - itinerary leg 6: return-leg airport (NULL if unused)
  fare_type         TEXT    - fare class/type code. [CONFIRM your codes]
  currency_of_sale  TEXT    - currency the ticket was sold in, e.g. 'NGN', 'USD'
  fare_usd          DOUBLE  - ticket fare in US Dollars (NULL if not applicable)
  fare_ngn          DOUBLE  - ticket fare in Nigerian Naira (NULL if not applicable)
  sales_levy_usd    DOUBLE  - sales levy in US Dollars
  sales_levy_ngn    DOUBLE  - sales levy in Nigerian Naira
  free_ticket       INTEGER - 1 if the ticket was free, otherwise 0

Notes for writing SQL:
  - A NULL leg means that leg was NOT used; it is not missing data.
  - Money columns are already numeric; use SUM/AVG directly.
  - Dates are real DATE values; filter with e.g.
    flight_date >= DATE '2023-01-01' AND flight_date < DATE '2023-04-01' for Q1 2023.
"""


def get_schema_prompt() -> str:
    """Return the schema description to embed in the LLM prompt (used in Part 2)."""
    return SCHEMA_TEXT.strip()
```

Notice the `[CONFIRM]` markers. `TSC`, the `pax_type` letter codes, and `fare_type` are business‑specific — I've given sensible placeholders, but **you should replace them with your organization's exact definitions.** This is the highest‑leverage 15 minutes you can spend on the whole project: a precise schema is the difference between DaveAI writing `WHERE pax_type = 'A'` correctly and it inventing a filter that returns garbage.

---

## 6. Prove the data actually loaded

Never trust a pipeline you haven't verified. This script counts the rows and runs a **real "requires‑math" query** — the exact kind DaveAI will generate later — so you can see the database doing arithmetic on properly‑typed columns.

Create `scripts/verify_load.py`:

```python
"""verify_load.py -- proves the data loaded and that the DB does the arithmetic."""
import sys
from pathlib import Path

# Let this script (in scripts/) import config and modules from src/
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

import duckdb
import config

con = duckdb.connect(str(config.DB_PATH), read_only=True)

print("Tables:", [r[0] for r in con.execute("SHOW TABLES").fetchall()])
print("manifest_reports rows:",
      con.execute("SELECT COUNT(*) FROM manifest_reports").fetchone()[0])
print("passenger_records rows:",
      con.execute("SELECT COUNT(*) FROM passenger_records").fetchone()[0])

# A genuine 'requires math' question, computed by the database:
# "How many flights in Q1 2025, and what was the total naira TSC and average pax?"
query = """
SELECT COUNT(*)               AS flights_q1_2025,
       SUM(tsc_amount_ngn)    AS total_tsc_ngn,
       ROUND(AVG(pax_count),1) AS avg_pax
FROM manifest_reports
WHERE flight_date >= DATE '2025-01-01'
  AND flight_date <  DATE '2025-04-01'
"""
print("Q1-2025 aggregate:", con.execute(query).fetchdf().to_dict("records"))
con.close()
```

Run it:

```powershell
python scripts\verify_load.py
```

You'll see the table names, the row counts for your two files, and a computed Q1‑2025 aggregate. The important thing conceptually: **that sum and average were computed by DuckDB from real numeric columns** — no LLM, no guessing. That's the deterministic core working, before the AI is even involved.

> **Prefer a notebook?** The same three checks work pasted into a Jupyter cell. Launch Jupyter with `jupyter notebook` (it's in `requirements.txt`), create `notebooks/01_verify_load.ipynb`, and put the `import`s + queries in a cell. Notebooks are handy for exploring your data interactively before Part 2.

---

## 7. `requirements.txt` (pinned)

Create this file in the project root. Versions are pinned to current, mutually‑compatible releases so your environment is reproducible.

```text
# --- Data pipeline (Part 1) ---
pandas==2.3.3
openpyxl==3.1.5
duckdb==1.4.5

# --- LLM client (used from Part 2 onward; safe to install now) ---
ollama==0.6.2

# --- Visualization (used in Parts 2/3; pinned now so deps are locked once) ---
plotly==6.9.0
matplotlib==3.10.9

# --- Notebooks (optional, for interactive verification) ---
jupyter==1.1.1
```

A few notes: `pandas` is pinned to the stable 2.3 line for maximum tutorial/compatibility comfort (the 3.x line also works with this code if you prefer it). `duckdb` is pinned to the 1.4 long‑term‑support line. The `ollama` here is the **Python client library**, which is separate from the Ollama *application* you installed in §2.3 — you don't pin the application in `requirements.txt`.

---

## 8. Part 1 recap — and what's coming next

**What you built:**
- An isolated Python environment with every dependency pinned.
- Ollama installed and running Qwen2.5‑7B‑Instruct locally on your Intel Arc GPU via Vulkan — verified, not assumed, to be using the GPU.
- A private GitHub repo (`dave-ai`) with a `.gitignore` that keeps raw flight data and the database off GitHub, and a README that explains the DaveAI name and architecture.
- A folder structure pre‑shaped for Parts 2 and 3.
- A tested data pipeline: `excel_loader.py` cleans both Excel files (currency strings → real numbers, two date formats parsed correctly, empty legs → true nulls, numeric column names fixed), `db_loader.py` loads them into two DuckDB tables, and `schema.py` documents every column in plain English for the LLM.
- A verifier proving the data loaded *and* that the database — not an LLM — computes sums and averages.

**What works right now if you stopped here:** you have a fully queryable local flight database. You can already answer any question by writing SQL yourself against `daveai.duckdb` (e.g. in `verify_load.py` or a notebook), and you can confirm your local LLM runs on the GPU. What you *can't* do yet is ask questions in plain English — that's the piece Part 2 adds.

**What Part 1 deliberately does *not* cover yet:**
- The text‑to‑SQL brain (turning English into SQL and running it) — **Part 2**.
- MLflow experiment tracking — **Part 2**.
- Voice input/output in English, Yoruba, Igbo, and Hausa — **Part 3**.
- A user interface and deployment — **Part 3**.

**Preview of Part 2 — the AI core + MLflow.** Part 2 is where DaveAI becomes conversational. You'll build the prompt that feeds `schema.py`'s dictionary plus your question into Qwen2.5 via the `ollama` Python client, get back a SQL query, safely execute it against DuckDB, and turn the raw result into a friendly English answer — closing the loop from the diagram in §0. Because an LLM's SQL quality depends heavily on how you prompt it, you'll wire in **MLflow** (writing to that `mlruns/` folder you already created) to systematically track each prompt variation, model setting, and accuracy score — so you can measure which approach actually works instead of guessing. Everything Part 2 needs is already in place from Part 1.
