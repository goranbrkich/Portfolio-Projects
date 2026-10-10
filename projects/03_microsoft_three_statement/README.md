# Microsoft Three Statement Model

Portfolio project 03 by Goran Brkich. Historical FY2017–FY2026 and forecast FY2027–FY2031. Information cutoff: **8 October 2026**. Microsoft fiscal years end on 30 June. Amounts are USD millions unless a label specifies another unit.

## Start here

1. Open `deliverables/Microsoft_Three_Statement_Model.xlsx` in Excel and start with **Overview**.
2. Read `deliverables/Microsoft_Three_Statement_Model_Report.docx` for the methodology, results and modeling limits.
3. On **Assumptions**, set **D4** to **1 Base**, **2 Upside** or **3 Downside**. Forecast years are in **O:S**. Each driver has an active row followed by three editable scenario rows.
4. Review the three linked statements, supporting schedules and **Audit**. Expanded native editable text boxes and live observations explain the calculations. Four output views cover Equity Research, Investment Banking, Private Equity and FP&A.
5. Open **Financial Modeling** for historical OLS, HC3 standard errors, confidence intervals, residuals and leverage. The complete Python code is in **C160:C332** and in `src/regression_analysis.py`.

Blue numbers are editable inputs, green working-sheet numbers link to another sheet, and black numbers are local calculations or summary results. Historical actuals remain unchanged when the forecast case changes. Set calculation to automatic in Excel, or recalculate and save after editing inputs.

The revision dated **10 October 2026** adds **35 native charts**, bringing the workbook to **37 charts on 21 worksheets**, **32 expanded analytical panels**, and **8,600 cell Notes**. Hover over a cell with a Note indicator to see definitions, calculation relationships and the exact formula. In current Excel versions these are **Notes**, available under **Review > Notes**. The operating charts and analytical observations follow the existing scenario selector. Historical-only charts retain actual data. Segment charts begin FY2023 on the current segment definition. Chart axes show USD billions while the underlying cells retain USD millions.

The compatibility update on **10 October 2026** replaces `TEXT` format-mask expressions in **52 live observations** with `FIXED` and numeric decimal arguments. This avoids interpreting English decimal and thousands masks when Excel recalculates under another regional setting. The 52 existing hover Notes show the updated formulas. Table rows have a common **24 pt** height on all 21 worksheets; titles and the separate Python code block keep their existing spacing. Financial inputs, statement formulas, regression calculations and the Word report are unchanged.

All 52 observations were recalculated in Base, Upside and Downside, with **165** forecast reconciliation checks. The saved file passes **1,120** independent statement comparisons, common-height checks for **2,558** rows and preservation checks for **10,392** untouched populated cells. A disposable LibreOffice Calc check clears the formula caches before reopening and recalculation. Microsoft desktop Excel is unavailable in the verification environment. See `evidence/excel_compatibility_verification.json`, `evidence/excel_compatibility_scenarios.json` and `evidence/excel_compatibility_feature_verification.json`.

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
| `src` | Source retrieval, normalization, numerical model, OLS/HC3 analysis, report builders and saved-workbook verification |
| `evidence` | Reconciliation results, original Excel checks, workbook row map, enhancement checks, regression CSV/JSON and figures |
| `tools` | Workbook-authoring utilities for native charts, Notes and expanded drawing text; authoring uses the Codex primary runtime |

## Offline numerical reproduction

Python 3.10 or later is sufficient for the forecast and saved-workbook checks. These scripts use the standard library.

```bash
python src/model.py --case Base
python src/verify_workbook.py
python src/verify_excel_compatibility.py
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

## Historical regression and Python reproduction

**Financial Modeling** calculates `ln(CFO) = intercept + slope * ln(revenue)` on FY2017–FY2026 actuals, then compares it with a regression of annual log changes on FY2018–FY2026. Excel formulas expose fitted values, residuals, leverage, HC3 weights, standard errors, two-sided p values and 95% confidence intervals. HC3 covariance retains the OLS coefficients. Student t inference uses `n-2` degrees of freedom as an explicit small-sample approximation. The visible reference table contains the 97.5th percentile for degrees of freedom 1–30.

Copy **C160:C332** into `regression_analysis.py` in VSCode, or open the supplied `src/regression_analysis.py` directly. Install NumPy and SciPy for numerical results; Matplotlib adds the comparison plot:

```bash
python -m pip install numpy scipy matplotlib
python src/regression_analysis.py --out evidence/regression
python src/verify_enhancements.py --run-copied-code
```

The standalone script embeds the ten original revenue and CFO observations and implements HC3 matrix covariance directly. It produces coefficient and diagnostic CSV files, a JSON result and the actual-versus-fitted CFO chart. It can instead read `--csv input.csv`, with `year,revenue,cfo` columns and positive values for consecutive fiscal years. Excel edits do not automatically rewrite the embedded Python observations.

The saved Excel results pass **56 regression summary comparisons** and **114 observation-level comparisons** against Python. The complete source extracted from worksheet cells was executed and reproduced the results. The preceding workbook's **8,424 populated cells**, including **4,515 formulas**, remain unchanged. Its original sheet order, scenario validation, freeze panes and two Overview charts are preserved. The revised saved workbook also passes **1,120** historical and active-case statement comparisons.

A disposable calculation checked **805 chart-helper values per scenario** in Base, Upside and Downside, including a zero-growth input. It also verified a later-year capex change and propagation of source errors into chart data. See `evidence/enhancement_scenario_verification.json`. The delivered workbook remains saved on Base.

These regressions describe historical association. Trending levels can overstate fit; the growth specification has only nine observations and a high-leverage fiscal year. HC3 does not correct serial correlation or structural changes. `EXP(fitted ln CFO)` is a fitted conditional median, without a mean-bias correction. The regression does not feed or replace the integrated operating forecast. Word sections 13 and 14 explain the results and reproduction steps.

## Primary source reproduction

Source extracts are included and the forecast runs offline. Full original annual-report HTML and SEC API responses are retrieved when a complete normalization rerun is needed:

```bash
python -m pip install -r requirements.txt
python src/fetch_sources.py
python src/normalize_filings.py
python src/model.py
python src/regression_analysis.py --out evidence/regression
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
