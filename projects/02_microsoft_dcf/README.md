# Microsoft DCF and Reverse DCF

Completed portfolio project 02 by Goran Brkich. Valuation cutoff **8 October 2026**. Historical context FY2017–FY2026. Five explicit forecasts **FY2027–FY2031**.

Base value **USD 286.35/share**, WACC 9.814%, observed close USD 522.61. Upside USD 390.99; downside USD 169.91. Values are conditional on visible assumptions.

## Deliverables

- [Excel model](deliverables/Microsoft_DCF_ReverseDCF_Model.xlsx): 34 sheets, live formulas, native What-If tables, 24 pt table rows, expanded analysis and native hover Notes.
- [Word report](deliverables/Microsoft_DCF_ReverseDCF_Report.docx): valuation, twelve reverse criteria, OLS/HC3/HAC, simulation and evidence.
- [Source evidence guide](deliverables/Source_Evidence_Guide.pdf) and [market review extracts](deliverables/Market_Source_Extracts.pdf).
- [Source crosswalk](docs/source_crosswalk.csv): 402 records, all 344 reported records highlighted, 51 prior disclosures unavailable and seven normalized absence records.
- Full [original annual reports](data/raw/reports/originals) and [highlighted review copies](data/raw/reports/highlighted), FY2017–FY2026, preserving native text and pages.

[Google Drive project folder](https://drive.google.com/drive/folders/18yihZ7JUvc994e_1v8JnLGtXcQ44O11K).

## Excel and Python

Start at Overview. DCF!E5 selects 1 Base, 2 Upside, 3 Downside. Edit blue assumptions, then review DCF, Sensitivity and each Reverse sheet. Automatic calculation including data tables is needed after edits. Both table axes replace the actual pricing inputs on the DCF sheet. The exported XLSX contains native dataTable records; one active forecast serves the primary valuation.

Python Code contains the complete plain-text program, also supplied as [src/valuation.py](src/valuation.py). Retain the complete folder structure in VSCode:

```bash
python -m pip install -r requirements.txt
python src/valuation.py
python src/valuation.py --case Upside
python src/valuation.py --excel deliverables/Microsoft_DCF_ReverseDCF_Model.xlsx
```

NumPy and SciPy implement OLS, HC3, Newey-West lag five, DCF, every reverse criterion and the identical 1,000 fixed simulation draws. Python reads mapped numeric drivers and latest historical financial cells directly, rather than trusting formula caches. Raw return history remains the frozen source baseline.

## Financial conventions

SBC stays expensed in EBIT. Economic capex includes new finance lease assets and excludes capitalized financing interest. NWC removes capex creditors and current lease financing. Existing finance lease liabilities and disclosed debt fair value enter the bridge; excess liquid assets and equity investments are added. Point shares plus the explicit existing-award proxy form the denominator. Funded buybacks consume equity value.

FY2027 cash flow is shortened by 265/365. Midyear discounting is the default, with a fiscal year-end alternative. Stable FCFF equals NOPAT times (1 - g / ROIC). Require WACC > g and ROIC > g.

Twelve reverse sheets cover growth, margins, capex, NWC, WACC, stable growth, ROIC, taxes, finance lease assets, normalized FCF, exit multiple and ownership. Invalid economic results are labelled. ROIC alone cannot reach the base market target; the no-reinvestment price ceiling is USD 326.91. Independent triangular simulation priors and exit multiples are explicit analyst assumptions; no peer or consensus sample is invented.

## Sources and validation

[Raw manifest](data/raw/source_manifest.json) and [market manifest](data/raw/market/market_source_manifest.json) record file hashes, URLs and retrieval times. The explicitly dated October author page selects ERP because the monthly archive has an inconsistent final date serial. Only observations at or before the cutoff enter processed data, even where original provider responses contain later metadata.

Audit compares independent scenario values, OLS/HC3/HAC statistics, reverse residuals, simulation and financial identities. Checks intentionally differ after edits; rerun Python for the edited workbook. Evidence includes formula, source, native XLSX and scenario behaviour checks. Desktop Microsoft Excel was not available for testing.

## Portable archive

The full ZIP is supplied for download and in Google Drive. GitHub download_parts contains the identical archive. After cloning, reassemble with:

```bash
python projects/02_microsoft_dcf/src/reassemble_archive.py --output Microsoft_DCF_ReverseDCF_Project02.zip
```

Part and archive SHA-256 values are verified. PACKAGE_MANIFEST.json inventories the portable project.
