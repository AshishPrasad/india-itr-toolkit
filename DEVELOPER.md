# Developer Guide

How to set up, run, test, and contribute to the **India ITR Toolkit**.

> **Privacy:** Follow the repo [privacy policy](README.md#privacy) — never put
> real financial data into code, tests, docs, or commit messages.

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

## 4. Run the tools

Each tool's invocation, flags, and input/output formats are documented in the
schedule's README — see [`ScheduleFA/README.md`](ScheduleFA/README.md) (the
`dividend-contribution` and `schedule-fa-csv` sections). Once installed, the
console scripts (`dividend-contribution`, `schedule-fa-csv`), the
`python -m <package>` form, and the `.cmd` wrappers all work; pass `--help` for
the full option list.

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

Each schedule folder is a self-contained project: one or more packages under
`src/`, a `tests/` suite (sample data generated into the gitignored
`tests/.tmp_data/`), optional `templates/`, plus `pyproject.toml`,
`requirements.txt`, `README.md`, `AGENTS.md`, and Windows `.cmd` launchers.

See the schedule's own README for the full, up-to-date file tree — e.g.
[`ScheduleFA/README.md` › Project layout](ScheduleFA/README.md#project-layout).

## Conventions

- Follow the coding and testing conventions in
  [`ScheduleFA/AGENTS.md`](ScheduleFA/AGENTS.md): modular `src/` code with pure
  core logic, I/O isolated in `readers.py`/`writer.py`, tolerant header
  detection, and allocation that reconciles to the total.
- Add or update tests for any behavior change; keep tests hermetic (synthetic
  data, `tmp_path` for outputs).
- New schedules go in their own top-level folder (see the repo conventions in
  the [root README](README.md#conventions)).

## Git workflow

```bash
git checkout -b <feature-branch>
# make changes; ensure `python -m pytest` passes
git commit
```

The repo's `.gitignore` (at the root) already excludes virtual environments,
caches, build artifacts, and Excel lock files.
