# Source materials and highlighted evidence

All original report bytes are retained under `data/raw`. Annotated copies are separate under `highlighted`. Native issuer PDFs retain their original page count and pagination. Complete SEC filing HTML is preserved where it supplies the selected observation, including issuer disclosures absent from companyfacts.

## Color legend

| Color | Financial meaning |
|---|---|
| Green | Income statement and earnings observations |
| Blue | Balance sheet, common equity, tangible book and deposit components |
| Yellow | Dividends, repurchases and cash distributions |
| Purple | Regulatory capital and note disclosures |

PDF bookmarks identify used pages. Annotation metadata names the input cell, bank, period, metric, value and source context. HTML hover text identifies the same source relationship.

## Find a model input

Use `data/processed/source_crosswalk.csv` or the Excel Source Crosswalk. The main value cells are in Historical Inputs column F. For Truist deposits, source component cells are in columns T and U on the same input row. Their total is calculated in the TFC financial build, rather than inserted into the raw value column.

The crosswalk retains the source period, metric, unit, filing date, accession, complete original path and URL. Exact XBRL anchors and verified table text identify the actual selected observation. Original USD and share values are scaled to millions in the model; ratios and per-share values remain unscaled.

Latest eligible comparative presentations may be selected from a later annual report, provided that report was public by 8 October 2026. Do not substitute an earlier presentation simply because the filename matches the metric year.

PNC's goodwill and other intangible assets are shown as a deduction in its non-GAAP reconciliation. The model records the positive asset magnitude, with the original negative reconciliation sign documented in selection metadata. The financial build subtracts goodwill once to derive other intangibles. This is distinct from assuming a missing disclosure is zero.

## Strength and limits of a match

An exact XBRL anchor or selected original table row is primary evidence. A PDF annotation matching the same bank, period, metric row and numeric value is secondary evidence; review the date column against the primary selected fact. Equal numbers alone never authorize a highlight. Multiple or ambiguous matches are refused and remain visibly unresolved.

Some native reports contain only rounded headline figures, or a preliminary capital percentage that differs from the final filed 10-Q. They cannot replace the exact filed selection. A PDF with no unambiguous match remains available as an original report, with its absence of annotations recorded in `qa/source_highlight_checks.json`.

Missing reported facts stay unavailable. Calculated provisions, reconstructed equity and share averages, analyst adjustments and FY2026 forecasts are labeled in the financial builds and their hover Notes. FY2026 annual actuals were unavailable at the frozen cutoff.

## Reproduction and preservation

`src/fetch_sources.py` retrieves primary SEC originals and unadjusted daily market responses. `src/extract_data.py`, `src/curate_inputs.py` and `src/complete_source_definitions.py` retain the selected fact provenance and explicit definition fixes. Live recollection is optional; the delivered analysis runs on the frozen files without credentials or downloads.

`src/highlight_sources.py` uses Poppler for word coordinates and pypdf for native PDF annotations, avoiding the previous native text-extraction failure. It verifies identical original and marked page counts and decoded page-content hashes. Full HTML text is compared before and after marking. `qa/source_highlight_checks.json` records every report result, highlight count and used page.

The archive includes source hashes and retrieval metadata. The frozen information cutoff remains 8 October 2026 even when the recovery download occurred later. Later filing observations are excluded from financial selection.
