# Developer Guide

How to set up, run, test, and contribute to the **India ITR Toolkit**.

> **Privacy:** Real workbooks may contain personal financial data. Never commit
> real `.xlsx` files or copy real values (amounts, holdings, totals, account
> details) into code, tests, docs, commit messages, or logs. Use the synthetic
> sample data generated under each tool's `tests/.tmp_data/` only.

## Prerequisites

- **Python 3.9+** (verified on 3.14) — `python --version`
- **git**
- **pip** (bundled with Python)

The repo is organised as one folder per ITR schedule (see the top-level
[`README.md`](README.md)). Each tool is self-contained with its own `src/`,
`tests/`, and packaging, so you set up and run each one from inside its folder.
The steps below use **`ScheduleFA`** as the example.

## 1. Clone

```bash
git clone <repository-url>
cd india-itr-toolkit/ScheduleFA
```

## 2. Create a virtual environment

Windows (PowerShell):

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

macOS / Linux:

```bash
python -m venv .venv
source .venv/bin/activate
```

## 3. Install

Editable install with dev dependencies (pytest):

```bash
pip install -e ".[dev]"
```

For runtime only (just `openpyxl`):

```bash
pip install -r requirements.txt
```

## 4. Run the tool

Any of these work once installed:

```bash
# Installed console script:
dividend-contribution --dividends dividends.xlsx --stocks ledger.xlsx --output out.xlsx

# Python module (no install needed if PYTHONPATH includes src/):
python -m dividend_contribution --dividends dividends.xlsx --stocks ledger.xlsx --output out.xlsx

# Windows convenience launcher (sets PYTHONPATH, no install needed):
dividend-contribution.cmd --dividends dividends.xlsx --stocks ledger.xlsx --output out.xlsx
```

See `--help` for all options, and `ScheduleFA/README.md` for input file formats.

## 5. Run the tests

```bash
python -m pytest
```

`pyproject.toml` sets `pythonpath = ["src"]`, so tests run without an install.
Sample workbooks are generated automatically before the suite runs. To
regenerate them manually:

```bash
python tests/generate_sample_data.py
```

## Project layout (per schedule)

```
ScheduleFA/
├── src/dividend_contribution/   # modular package (one responsibility per module)
├── tests/                       # pytest suite (+ generated tests/.tmp_data/, gitignored)
├── pyproject.toml               # packaging + pytest config (src layout)
├── requirements.txt             # runtime dependency (openpyxl)
├── dividend-contribution.cmd    # Windows launcher (batch script, not an .exe)
├── README.md                    # tool usage
└── AGENTS.md                    # coding & testing conventions
```

## Conventions

- Follow the coding and testing conventions in
  [`ScheduleFA/AGENTS.md`](ScheduleFA/AGENTS.md): modular `src/` code with pure
  core logic, I/O isolated in `readers.py`/`writer.py`, tolerant header
  detection, and allocation that reconciles to the total.
- Add or update tests for any behavior change; keep tests hermetic (synthetic
  data, `tmp_path` for outputs).
- New schedules go in their own top-level folder following the same structure.

## Git workflow

```bash
git checkout -b <feature-branch>
# make changes; ensure `python -m pytest` passes
git commit
```

The repo's `.gitignore` (at the root) already excludes virtual environments,
caches, build artifacts, and Excel lock files.
