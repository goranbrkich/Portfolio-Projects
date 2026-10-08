# Microsoft Three Statement Model

Portfolio project 03 by Goran Brkich. Historical FY2017–FY2026 and forecast FY2027–FY2031. Information cutoff: **8 October 2026**. Microsoft fiscal years end on 30 June. Amounts are USD millions unless a label specifies another unit.

## Start here

1. Open `deliverables/Microsoft_Three_Statement_Model.xlsx` in Excel and start with **Overview**.
2. Read `deliverables/Microsoft_Three_Statement_Model_Report.docx` for the methodology, results and modeling limits.
3. On **Assumptions**, set **D4** to **1 Base**, **2 Upside** or **3 Downside**. Forecast years are in **O:S**. Each driver has an active row followed by three editable scenario rows.
4. Review the three linked statements, supporting schedules and **Audit**. Native editable text boxes explain the calculations. Four output views cover Equity Research, Investment Banking, Private Equity and FP&A.

Blue numbers are editable inputs, green working-sheet numbers link to another sheet, and black numbers are local calculations or summary results. Historical actuals remain unchanged when the forecast case changes. Set calculation to automatic in Excel, or recalculate and save after editing inputs.

## Scope and interpretation

The Base case grows revenue to approximately **$648.1 billion** in FY2031. Operating margin falls as the depreciation burden of new infrastructure increases. Free cash flow depends materially on the assumed decline in cash capex intensity. Forecasts are research assumptions, not Microsoft guidance.

The model separates cash capex, capital expenditure payables, new finance lease assets, debt face value and carrying debt. PPE uses annual depreciation cohorts. It includes debt and lease payment schedules, working capital, taxes, equity and year-end liquidity funding. No unexplained asset or equity balancing entry is used.

DCF, Reverse DCF, comparable company valuation, LBO and an investment recommendation belong to separate portfolio cases. The author's investment thesis will be developed through an interview. The four professional views interpret this common operating model.

## Files

| Directory | Contents |
|---|---|
| `deliverables` | Editable Excel workbook and explanatory Word report |
| `data/raw` | Source manifest, extracted primary statement tables and selected SEC XBRL facts |
| `data/processed` | Historical statements with provenance, assumptions and three-case forecasts |
| `src` | Source retrieval, normalization, numerical model, report builder and saved-workbook verification |
| `evidence` | Reconciliation results, original Excel checks, workbook row map, saved-value checks and figures |
| `tools` | Utility for restoring native drawing text if a separate supported workbook-authoring workflow requires it |

## Offline numerical reproduction

Python 3.10 or later is sufficient for the forecast and saved-workbook checks. These scripts use the standard library.

```bash
python src/model.py --case Base
python src/verify_workbook.py
```

`model.py` recalculates **all three cases** every time. `--case` changes only the scenario printed to the terminal. It writes forecast JSON/CSV and validation results. To change Python drivers, edit `data/processed/assumptions.json`.

The delivered numerical implementation passes **90** forecast reconciliations across the three cases. Historical evidence contains **170** source and accounting controls. The original Excel verification contains **210** output comparisons across the three cases, an input-change check and a missing unselected-case input check.

### Check edited Excel inputs

Excel and Python input stores are separate. Recalculate and save in Excel before reading saved values.

```bash
python src/verify_workbook.py --use-excel-inputs
python src/export_excel_assumptions.py --output evidence/excel_assumptions.json
```

The verifier reads all saved scenario drivers, allocations and FY2032 tail assumptions. It compares the active Excel statements with an independent Python forecast, checks the historical figures and scans saved errors. It **does not execute Excel recalculation**. A stale saved value or different assumption set causes a comparison failure instead of silently changing the expected result.

To make exported Excel assumptions the Python model inputs, explicitly copy the exported JSON to `data/processed/assumptions.json`. Retain a backup of the original assumptions. The verifier supports `--workbook PATH` for another saved copy.

## Primary source reproduction

Source extracts are included and the forecast runs offline. Full original annual-report HTML and SEC API responses are retrieved when a complete normalization rerun is needed:

```bash
python -m pip install -r requirements.txt
python src/fetch_sources.py
python src/normalize_filings.py
python src/model.py
python src/build_report.py
```

Source retrieval requires internet access to Microsoft and the SEC, uses an identifying User-Agent and limits concurrent requests. The normalization applies the fixed filing-date cutoff. Later changes in source pages may alter their byte hashes. Compare with the original `data/raw/source_manifest.json` when refreshing.

`build_report.py` requires the supplied `evidence/excel_verification.json` and produces the Word report and its figures from the datasets. It does not rebuild the Excel workbook. The workbook is supplied as an editable native-formula deliverable.

## Portable archive

The downloadable `Microsoft_Three_Statement_Project03.zip` contains the Excel and Word deliverables, source extracts, model code, assumptions, instructions and verification evidence. Its `PACKAGE_MANIFEST.json` records a SHA-256 hash for each project file. To create a fresh archive after editing the project:

```bash
python src/package_project.py --output Microsoft_Three_Statement_Project03.zip
```

The packaging script excludes Python caches and existing ZIP archives, then verifies the archive CRC and each included file hash. The full original annual-report HTML is retrieved separately with `fetch_sources.py`.

## Data comparability and model limits

- FY2016 is opening-balance context only. It is not an eleventh historical year.
- Each historical statement uses the latest available comparative presentation by the cutoff. The CSV retains source report year, source table index and URL.
- Segment definitions change. FY2023–FY2026 use the current definition; FY2023 cross-definition growth is unavailable.
- Note disclosures can be rounded. The cash-flow line “depreciation, amortization and other” is not identical to the sum of the separate note disclosures.
- D&A and operating lease allocations by segment are assumptions. New operating leases use an annual approximation, not a contract-level amortization engine.
- Tax balances use planning ratios. Diluted shares are an independent driver and do not automatically follow repurchase cash.
- Cash capex includes modeled capitalized cash interest. Finance lease principal is a separate financing outflow.
- No unannounced acquisitions, asset disposals, impairment events or currency changes are forecast. Goodwill, land and AOCI remain constant.
- A minimum cash funding requirement is not a committed credit facility. Large ending cash balances depend on the reinvestment and payout assumptions.

## Primary references

- [Microsoft FY2026 Form 10-K](https://www.sec.gov/Archives/edgar/data/789019/000119312526323660/msft-20260630.htm), filed 29 July 2026.
- [Microsoft annual reports](https://www.microsoft.com/en-us/Investor/annual-reports.aspx).
- [SEC Company Facts for Microsoft](https://data.sec.gov/api/xbrl/companyfacts/CIK0000789019.json).

See the source manifest and historical CSV for the exact annual-report URL and table used for each observation. All deliverables are in English.
