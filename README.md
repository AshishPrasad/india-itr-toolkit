# India ITR Toolkit

A collection of small, self-contained tools to help prepare **Income Tax Return
(ITR)** filings in India. Each ITR schedule lives in its own top-level folder so
the tools can evolve independently while sharing a common home.

## Privacy

Real workbooks may contain personal financial data. Do **not** copy real values
(amounts, holdings, totals, account details) into source code, tests,
documentation, commit messages, or logs. Tests must use synthetic sample data
only.

## Schedules

| Folder                     | Schedule                | Status      | Description                                                                 |
| -------------------------- | ----------------------- | ----------- | --------------------------------------------------------------------------- |
| [`ScheduleFA/`](ScheduleFA/) | Schedule FA (Foreign Assets) | ✅ Available | Per-lot dividend contribution, plus CSV generation for Schedule FA section A3. |
| `ScheduleCG/`              | Schedule CG (Capital Gains)  | 🚧 Planned  | Capital-gains computation for equity/foreign holdings.                      |
| `ScheduleOS/`              | Schedule OS (Other Sources)  | 🚧 Planned  | Interest, dividends and other income.                                       |

See each folder's own `README.md` for purpose, input formats, usage and tests.
For repo setup, see [`DEVELOPER.md`](DEVELOPER.md).

## Conventions

- One schedule per top-level folder; keep each tool's code, tests and docs
  inside its folder.
- Repo name and folder names are **year-agnostic** — keep the assessment/financial
  year (e.g. `FY2025-26`) inside data files, not in folder names, so the tools
  serve every year.
- Each tool follows a `src/` + `tests/` layout with a `README.md`; see
  `ScheduleFA/AGENTS.md` for the coding and testing conventions.

## License

Released under the [MIT License](LICENSE).
