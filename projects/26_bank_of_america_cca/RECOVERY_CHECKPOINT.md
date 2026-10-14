> Historical recovery note. Superseded by the verified final delivery on 14 October 2026. See [README.md](README.md) for current files, results and reproduction instructions.

# 26 — Bank of America Comparable Company Analysis

**Delivery status: working artifacts created; final source annotation and publication interrupted. This is a recovery checkpoint, not a completed delivery.**

## Agreed scope

- Target: Bank of America (BAC).
- Nine peers: JPMorgan Chase (JPM), Citigroup (C), Wells Fargo (WFC), Goldman Sachs (GS), Morgan Stanley (MS), U.S. Bancorp (USB), PNC Financial Services (PNC), Truist (TFC), Capital One (COF).
- Frozen price and information cutoff: 8 October 2026.
- FY2017–FY2026 scope: FY2017–FY2025 reported actuals; FY2026 analyst earnings estimate, with H1 2025, H1 2026 and June 2026 LTM disclosed separately. FY2016 is an opening reference. FY2026 full-year actuals were unavailable at the cutoff.
- One formula-based Excel workbook with plain-text copyable Python, one Word report, and a full portable source archive.
- Preserve uniform 24 pt table rows, native hover Notes, explanatory charts, expanded interpretation, and complete raw sources plus separately highlighted copies.
- The withdrawn corporate Excel and Word reference styles remain withdrawn.

## Work recorded before interruption

- Working Excel: 43 worksheets, ten company financial histories, peer statistics, valuation, earnings forecast, profitability, credit, funding, capital, historical multiples, ranking, scalar OLS/HC3, influence, two native What-If sensitivity tables, credit stress, market risk and correlations.
- The source-input table contains 4,920 metric positions across 120 bank-period records. Missing information stays unavailable rather than becoming zero.
- Python reproduces CCA, explicit NumPy HC3 covariance, multivariate regression, seeded 5,000-replication pairs bootstrap, leave-one-out diagnostics and 252-observation price risk.
- Word report created and visually reviewed: 12 pages. It documents method, accounting differences, limitations and issuer-specific interpretation.
- Original collection retained locally: 20 SEC companyfacts/submissions JSON files; ten original daily price responses and ten processed price CSVs; 124 complete SEC filing HTML reports plus the PNC official financial supplement; 40 complete native issuer PDFs.
- Separate highlighted full HTML copies created, with original financial text preservation checked.
- Excel previews were reviewed for all 43 worksheets. Baseline Audit values matched Python to rounding; formula-error scan returned no standard errors.
- A native PDF word-extraction crash interrupted source highlighting. Subsequently the execution service reported that the environment was offline. No full final archive or file publication was confirmed.

## Frozen baseline results

All prices in USD per common share; these are conditional analytical estimates.

| Measure | Baseline |
|---|---:|
| BAC market close, 8 October 2026 | 53.61 |
| Peer-median weighted implied price | 55.8857053856517 |
| Regression conditional mean BAC price | 54.60114442287685 |
| Bivariate OLS slope | 15.606942636259697 |
| Bivariate OLS intercept | -0.5802749450539989 |
| HC3 slope standard error | 4.037544831943776 |
| Regression peer observations | 9 |
| Regression residual degrees of freedom | 7 |

Weights are 25% P/E, 25% P/B and 50% P/TBV. Market risk uses price returns excluding dividends and an equal-weight nine-peer bank benchmark, not the S&P 500.

## Recovery location and remaining work

Working project root:
`/workspace/scratch/f19534a460e1/bank_cca_20261008/Bank_of_America_CCA_Project26`

Created working deliverables:
- `deliverables/Bank_of_America_CCA_Model.xlsx`
- `deliverables/Bank_of_America_CCA_Report.docx`

Canonical ArtifactTool builder:
`/tmp/bank_cca26_ipppczh1/build_cca.mjs`

Source-marking script:
`/workspace/scratch/f19534a460e1/bank_cca_20261008/highlight_sources.py`

HTML marking checkpoint:
`qa/html_highlight_checkpoint.json`

Next steps:
1. Restore access to the existing working files; do not recollect the data or rebuild the report.
2. Replace the crashing MuPDF word extraction with Poppler `pdftotext -bbox`; retain original PDF page streams and create annotation-only marked copies. Resume from the HTML checkpoint.
3. Update the Excel Source Crosswalk with final matched PDF pages and annotation paths. Clarify provision residual and period-end credit-risk labels.
4. Verify the commercial-peer selector, forecast-driver changes, corrected numeric HC3 confidence values and native sensitivity sample points; restore the default inputs.
5. Package all complete original sources, marked copies, crosswalks, code and deliverables. Exclude internal parser caches and rendering intermediates.
6. Save the downloadable XLSX, DOCX and full archive, upload them to the project Drive folder, and publish the project code, source package and final delivery record here. Only then mark project 26 completed.

Project Drive folder already created:
https://drive.google.com/drive/folders/1Rf01Q04YXzMhAEJVBie6X4pG6GjrkTE7

The folder was created successfully; artifact upload was not confirmed. This checkpoint preserves the known state without claiming completed publication.

