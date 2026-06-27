# Dividend Contribution Tool

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

Requires Python 3.9+.

```bash
pip install -r requirements.txt        # just the runtime dependency (openpyxl)
# or, to install as a command and for development:
pip install -e .[dev]
```

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

## Project layout

```
.
├── src/dividend_contribution/   # modular package
│   ├── parsing.py               # value/date parsing helpers
│   ├── headers.py               # header detection & column resolution
│   ├── readers.py               # read dividend / ledger workbooks
│   ├── allocation.py            # core allocation maths
│   ├── computation.py           # orchestration (compute -> Result)
│   ├── writer.py                # write the result workbook
│   └── cli.py                   # argparse command-line interface
├── tests/                       # pytest suite + sample-data generator
│   └── generate_sample_data.py  # writes sample .xlsx to gitignored tests/.tmp_data/
├── pyproject.toml               # packaging + pytest config
├── requirements.txt
├── dividend-contribution.cmd    # Windows wrapper
└── AGENTS.md                    # conventions for agents/contributors
```

---

## Running the tests

```bash
python -m pytest
```

The suite covers parsing, header detection, allocation maths (including partial
sells and over-sell warnings), workbook reading, and an end-to-end CLI run that
verifies the output workbook and that contributions foot to the dividend total.
Sample workbooks are generated automatically before the tests run, into the
gitignored `tests/.tmp_data/` folder (they are not committed).
