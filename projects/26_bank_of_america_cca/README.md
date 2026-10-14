# Bank of America Comparable Company Analysis

Project 26 in the Finance Portfolio. Frozen valuation and information cutoff: **8 October 2026**.

The target is Bank of America. The nine peers are JPMorgan Chase, Citigroup, Wells Fargo, Goldman Sachs, Morgan Stanley, U.S. Bancorp, PNC, Truist and Capital One.

## Delivered files

- `deliverables/Bank_of_America_CCA_Model.xlsx`: 43 worksheets with native formulas, charts, 24 pt table rows and hover Notes, plus complete plain-text Python.
- `deliverables/Bank_of_America_CCA_Report.docx`: 12-page English analytical report.
- `data/raw`: complete unmodified original datasets and reports.
- `highlighted`: separate complete marked reports with original pagination.
- `data/processed/source_crosswalk.csv`: searchable cell-to-source mapping.
- `docs/SOURCE_GUIDE.md`: annotation legend, source conventions and unresolved-match interpretation.
- `qa`: independent numerical, source-preservation and native-feature evidence.

The full portable archive contains all these folders. The GitHub copy uses numbered archive parts because individual Git blobs have a size limit. Run `python reassemble_archive.py` after downloading all parts in `archive_parts` to recover the same ZIP and verify its SHA-256.

[Google Drive project folder](https://drive.google.com/drive/folders/1Rf01Q04YXzMhAEJVBie6X4pG6GjrkTE7).

## Financial scope

The requested FY2017–FY2026 horizon comprises FY2017–FY2025 reported actuals and an explicitly labeled FY2026 analyst earnings estimate. December-year-end banks had no complete FY2026 actual statements at the cutoff. H1 2025, H1 2026 and June 2026 LTM are separate. FY2016 supplies opening balances only.

Valuation uses P/E, P/B and P/TBV on common equity. The default implied price combines 25%, 25% and 50% weights. Deposits and funding costs are operating features of a bank; industrial EV/EBITDA is not the primary convention.

Reported book and tangible-book measures take precedence over straight subtraction where available. Preferred par values are not treated as carrying amounts. Reported and reconstructed bases remain visible. The provision fallback is explicitly labeled as a pretax bridge residual.

Truist deposit totals are calculated in its financial build from separately reported noninterest-bearing and interest-bearing liabilities. The original components are retained beside the corresponding source-input rows in columns T and U. Gross and net loan definitions remain separate.

## Reproduce the analysis

Extract the full project directory and install the dependencies in `requirements.txt`.

```bash
python src/analysis.py --cohort 1
python src/analysis.py --cohort 2
python src/analysis.py --from-excel deliverables/Bank_of_America_CCA_Model.xlsx
```

The default uses all nine peers. Cohort 2 excludes Goldman Sachs and Morgan Stanley. The saved-Excel option reads edited input values, prices, peer selection, weights and forecast assumptions without authoring a workbook. It writes a separate `python_results_from_excel.json`.

Python implements OLS HC3 covariance explicitly, multivariate regression, a 5,000-replication pairs bootstrap with seed 20261008, leave-one-out estimates, forecast earnings, credit stress, book-value sensitivities and market risk. Market statistics use 252 price-return observations; dividends are excluded. BAC beta uses an equal-weight nine-peer bank benchmark rather than the S&P 500.

## Verification and limits

Native Excel and Python agree on valuation, earnings forecast, bivariate coefficients, HC3 standard errors and conditional-mean confidence bounds. The seven-peer selector and changed revenue driver were checked while preserving historical values. Two native What-If Data Tables were checked at their base and non-base points.

ArtifactTool calculation and a disposable LibreOffice recalculation were tested. Microsoft Excel desktop was unavailable. The final workbook preserves the native Notes, charts and What-If features from the canonical export; the LibreOffice test copy is not delivered.

The original text of marked HTML and the decoded page-content streams of marked PDFs were preserved. Exact XBRL anchors and selected original rows are the strongest source links. Numeric PDF evidence is matched to metric rows, with ambiguous matches explicitly flagged. The source guide explains secondary PDF evidence and date-column checks.

Historical fiscal-year multiples are ex post comparisons, not a tradable information-availability backtest. Small peer samples, accounting differences and acquisition effects limit the precision of valuation and regression estimates.
