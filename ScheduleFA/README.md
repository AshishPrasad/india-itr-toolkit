# Schedule FA Tools

Helpers for the **Schedule FA (Foreign Assets)** part of an Indian ITR. This
folder contains two command-line tools:

1. **`dividend-contribution`** — computes how much of a year's dividends each
   stock lot contributed (per-lot dividend attribution).
2. **`schedule-fa-csv`** — generates the CSV for **section A3 of Schedule FA**
   (Table A3 — *financial interest in any entity*) from a per-lot Excel summary.
   See [Schedule FA CSV generator (section A3)](#schedule-fa-csv-generator-section-a3)
   below.

---

## Dividend Contribution Tool

A command-line tool that computes **how much of a year's dividends each stock lot
contributed**, based on the quantity of each lot actually held on every dividend
payout date. It is designed for Indian ITR / Schedule FA workflows where foreign
shares need per-lot dividend attribution.

Given two Excel files — a **dividend schedule** and a **stock-transaction
ledger** (buys *and* sells, including partial sells) — it produces an Excel file
with the dividend contribution of each lot, plus a grand total.

---

## Reporting period: calendar year (not the April–March financial year)

Schedule FA is reported on a **calendar-year basis** — the year ending on
**31 December** of the relevant previous year — **not** the Indian April–March
financial year. So for **AY 2025-26 (FY 2024-25)**, Schedule FA covers
**1 January 2024 to 31 December 2024**, and foreign assets must be disclosed if
held anytime in that window, even if disposed of before 31 March.

Accordingly, this tool and its dividend/stock inputs are organised by calendar
year: dividend payout dates, **peak value** (the maximum during the calendar
year) and **closing value** (as on 31 December) all use the 1 Jan–31 Dec period.

> Verify the current Schedule FA rules against the official Income Tax
> Department guidance and ITR instructions for your assessment year before
> filing, as requirements can change.

---

## How allocation works

For each dividend (payout `date`, amount `A` in INR):

1. The **net held quantity** of every lot on that date is computed as
   `sum(Buys with date ≤ payout date) − sum(Sells with date ≤ payout date)`.
2. `A` is split across the lots **in proportion to their positive net held
   quantity**.

So a lot earns dividends only for the period it was held: lots bought after a
payout date earn nothing for it, and a partially-sold lot earns proportionally
less afterwards. The per-lot contributions always sum back to the total dividend
(largest-remainder rounding keeps the output column footing exactly).

> Eligibility uses the **payout date** in the dividend file as the cut-off.
> If you need record/ex-dividend-date precision, put those dates in the
> dividend file instead.

---

## Installation

Requires Python 3.9+ (runtime dependency: `openpyxl`). For full setup — virtual
environment, editable install, and dependencies — see
[`DEVELOPER.md`](../DEVELOPER.md). The `python -m <package>` and `.cmd` forms
below also run without installing.

---

## Input file formats

> **Dates:** Native Excel date cells are read directly. Dates stored as **text**
> are interpreted as **day-first** (`DD/MM/YYYY` or `DD-MM-YYYY`), matching the
> Indian locale — e.g. `04/03/2025` is **4 March 2025**, not 3 April. ISO
> `YYYY-MM-DD` and `DD-Mon-YYYY` are also accepted.

### 1. Dividends file

One row per dividend payout. Default column headers (override with CLI flags):

| Column                      | Meaning                                  |
| --------------------------- | ---------------------------------------- |
| `Date`                      | Dividend payout date                     |
| `Total Dividend (USD)`      | Gross dividend paid on that date, in USD |
| `USD -> INR (TT Buy SBI)`   | USD→INR exchange rate for that date      |

The tool computes `INR = USD × rate` for each row. Extra columns are ignored.
Leading title rows, blank rows and a trailing `Total` row are handled
automatically. Currency/number text is parsed transparently, including the `₹`
symbol and **Indian lakh-style digit grouping** (e.g. `₹12,34,567.89`) as well as
standard grouping (`₹1,234,567.89`).

> **TTBR (TT Buy Rate):** the SBI TT Buy reference rate for USD→INR, commonly
> used to convert foreign dividend income for Indian tax filing. Historical
> daily rates can be sourced from
> [`sahilgupta/sbi-fx-ratekeeper`](https://github.com/sahilgupta/sbi-fx-ratekeeper).

Worked example (`dividends.xlsx`):

| Date       | Total Dividend (USD) | USD -> INR (TT Buy SBI) |
| ---------- | -------------------- | ----------------------- |
| 2025-02-15 | 100.00               | ₹80.00                  |
| 2025-05-20 | 100.00               | ₹85.00                  |
| 2025-08-20 | 100.00               | ₹90.00                  |
| 2025-11-20 | 100.00               | ₹95.00                  |

The tool computes `INR = USD × rate` per row → ₹8,000 + ₹8,500 + ₹9,000 + ₹9,500
= **₹35,000** total dividend to allocate.

### 2. Stock ledger file

One row per transaction. Default column headers:

| Column       | Meaning                                             |
| ------------ | --------------------------------------------------- |
| `Date`       | Transaction date                                    |
| `Buy/Sell`   | `Buy` or `Sell`                                     |
| `Lot number` | Identifier of the lot (a `Sell` reduces this lot)   |
| `Quantity`   | Number of shares (positive); a `Sell` may be partial |

A `Sell` points at the lot it reduces and may be smaller than the lot
(partial sell).

Worked example (`ledger.xlsx`):

| Date       | Buy/Sell | Lot number | Quantity |
| ---------- | -------- | ---------- | -------- |
| 2025-01-01 | Buy      | 1          | 10       |
| 2025-05-01 | Buy      | 2          | 10       |
| 2025-08-01 | Sell     | 1          | 5        |

Lot 1 is held in full for the first two dividends, then halved (5 of 10 sold)
before the last two; lot 2 is bought just before the second dividend. Running
this against the dividend example above yields the [output](#output) below.

---

## Usage

```bash
# Using the installed console script:
dividend-contribution --dividends dividends.xlsx --stocks ledger.xlsx --output result.xlsx

# Or via the Python module (no install needed):
python -m dividend_contribution --dividends dividends.xlsx --stocks ledger.xlsx --output result.xlsx

# Or via the Windows wrapper:
dividend-contribution.cmd --dividends dividends.xlsx --stocks ledger.xlsx --output result.xlsx
```

### Column-name overrides

If your files use different headers (auto-detection covers common aliases, but
you can be explicit):

```bash
python -m dividend_contribution \
  --dividends div.xlsx --stocks ledger.xlsx --output out.xlsx \
  --div-usd-col "Gross Dividend (USD)" --div-rate-col "FX Rate" \
  --stk-type-col "Action" --stk-lot-col "Lot" --stk-qty-col "Shares"
```

Run `python -m dividend_contribution --help` for the full option list, including
`--div-sheet` / `--stk-sheet` to select a specific worksheet.

### Output

An Excel workbook with one row per lot followed by a bold total row. Columns:

| Column                        | Meaning                                                        |
| ----------------------------- | -------------------------------------------------------------- |
| `Lot number`                  | Lot identifier carried over from the stock ledger              |
| `Net Held Qty`                | Shares still held for the lot (total Buys − total Sells)       |
| `Dividend Contribution (INR)` | Dividend the lot earned across all payout dates it was held    |

The bold `Total` row sums `Net Held Qty` and `Dividend Contribution (INR)`; the
contribution total equals the dividend allocated across all lots.

Worked example (the result of running the dividend and ledger examples above):

| Lot number | Net Held Qty | Dividend Contribution (INR) |
| ---------- | ------------ | --------------------------- |
| 1          | 5.0000       | ₹18,416.67                  |
| 2          | 10.0000      | ₹16,583.33                  |
| **Total**  | **15.0000**  | **₹35,000.00**              |

> **Note:** Lot labels that begin with `=`, `+`, `-`, `@`, or a tab/carriage
> return are written with a leading apostrophe (e.g. `'=ABC`) so spreadsheets
> render them as text rather than formulas — a safeguard against spreadsheet
> formula injection if you share the generated file.

---

## Schedule FA CSV generator (section A3)

The `schedule-fa-csv` tool generates the CSV for **section A3 of Schedule FA**
(Table A3 — *financial interest in any entity*). It converts a per-lot Excel
summary into the exact 12 columns that section A3 expects. A natural source for
the per-lot peak/closing/dividend figures
is the output of the dividend-contribution tool, enriched with peak/closing
values and entity details.

### Input file

One row per lot/holding. Default column headers (override with `--*-col` flags;
common aliases are auto-detected):

| Column                  | Meaning                                              |
| ----------------------- | ---------------------------------------------------- |
| `Date`                  | Date the interest/holding was acquired               |
| `Lot number`            | Lot identifier (for reference)                       |
| `Initial value`         | Cost/initial value of the investment (INR)           |
| `Peak value in INR`     | Peak value during the period (INR)                   |
| `Closing value in INR`  | Closing value at period end (INR)                    |
| `Dividend Contribution` | Gross amount paid/credited during the period (INR)   |
| `Sale proceeds`         | Gross proceeds from sale/redemption (INR); **optional**, defaults to 0 |
| `Country/Region name`   | Country/Region name                                  |
| `Country Name and Code` | Country name and ITR country code                    |
| `Name of entity`        | Name of the entity                                   |
| `Address of entity`     | Address of the entity                                |
| `ZIP Code`              | ZIP/PIN code (text; leading zeros preserved)         |
| `Nature of entity`      | Nature of the entity                                 |

A `Quantity` column may be present but is ignored. Currency text (e.g.
`₹1,00,000.00`) and lakh-style grouping are parsed transparently.

> **Use the exact headers (or the template).** Headers are matched by exact name
> first, then by alias. Aliases prefer the most specific match, but with unusual
> headers auto-detection can still pick the wrong column. For non-standard sheets,
> use the [input template](templates/) or pass explicit `--*-col` flags (e.g.
> `--peak-col "Maximum INR"`). If two fields would map to the same column the tool
> stops with an *ambiguous header mapping* error rather than emit wrong figures.

> **Templates:** ready-to-use header files live in [`templates/`](templates/):
> `fa_input_template.xlsx` (empty input sheet with the expected headers, ZIP
> column pre-formatted as text) and `fa_output_headers.csv` (the exact output
> columns). **Copy the input template out of the repo before entering real data**
> so personal information never lands in git. Regenerate with
> `python templates/generate_templates.py`.

### Output CSV

One row per holding, with these columns in order:

```
Country/Region name, Country Name and Code, Name of entity, Address of entity,
ZIP Code, Nature of entity, Date of acquiring the interest,
Initial value of the investment, Peak value of investment during the Period,
Closing balance,
Total gross amount paid/credited with respect to the holding during the period,
Total gross proceeds from sale or redemption of investment during the period
```

Amounts are written as **whole rupees** (no `₹` symbol or thousands separators,
so the file parses cleanly as CSV). The acquisition date defaults to `YYYY-MM-DD`
(configurable with `--date-format`). The file is UTF-8 (with BOM) for clean
opening in Excel, and text fields are guarded against spreadsheet formula
injection.

> **Rounding (no drift):** each numeric column is rounded to whole rupees using
> **largest-remainder (Hamilton) rounding** — every value is floored, then the
> rupees lost to flooring are handed back to the rows with the largest fractional
> parts. This guarantees the column sums **exactly** to the column total rounded
> to the nearest rupee, so the reported figures never drift away from the source
> total the way independent per-row rounding would.
>
> *Example:* three lots of ₹33.34, ₹33.33, ₹33.33 (total ₹100) are written as
> **34, 33, 33** — summing to 100. Rounding each value independently would give
> 33 + 33 + 33 = **99**, losing a rupee against the ₹100 total.

### Usage

```bash
# Installed console script:
schedule-fa-csv --input holdings.xlsx --output schedule_fa.csv

# Python module (no install needed):
python -m schedule_fa_csv --input holdings.xlsx --output schedule_fa.csv

# Windows wrapper:
schedule-fa-csv.cmd --input holdings.xlsx --output schedule_fa.csv
```

Run `python -m schedule_fa_csv --help` for all options, including `--sheet`,
`--date-format`, and per-column override flags (e.g. `--peak-col`, `--zip-col`).

---

## Project layout

```
.
├── src/dividend_contribution/   # dividend-contribution tool
│   ├── parsing.py               # value/date parsing helpers
│   ├── headers.py               # header detection & column resolution
│   ├── readers.py               # read dividend / ledger workbooks
│   ├── allocation.py            # core allocation maths
│   ├── computation.py           # orchestration (compute -> Result)
│   ├── writer.py                # write the result workbook
│   └── cli.py                   # argparse command-line interface
├── src/schedule_fa_csv/         # schedule-fa-csv tool (Schedule FA section A3 CSV)
│   ├── reader.py                # read the per-lot input workbook
│   ├── writer.py                # write the 12-column Schedule FA CSV
│   └── cli.py                   # argparse command-line interface
├── templates/                   # header-only input/output templates (safe to copy)
│   └── generate_templates.py    # regenerates the template files
├── tests/                       # pytest suite + sample-data generator
│   └── generate_sample_data.py  # writes sample .xlsx to gitignored tests/.tmp_data/
├── pyproject.toml               # packaging + pytest config
├── requirements.txt
├── dividend-contribution.cmd    # Windows wrapper (dividend tool)
├── schedule-fa-csv.cmd          # Windows wrapper (CSV generator)
└── AGENTS.md                    # conventions for agents/contributors
```

---

## Running the tests

Run the suite with `python -m pytest` — see
[`DEVELOPER.md`](../DEVELOPER.md#5-run-the-tests) for details. It covers parsing,
header detection, allocation maths (including partial sells and over-sell
warnings), workbook reading, CSV generation, and end-to-end CLI runs that verify
outputs foot to their totals. Sample workbooks are generated automatically
(into the gitignored `tests/.tmp_data/`) before the tests run.
