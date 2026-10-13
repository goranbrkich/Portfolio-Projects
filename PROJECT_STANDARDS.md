# Portfolio Project Standards

Confirmed by Goran Brkich on 12 October 2026. Apply these requirements to future similar financial and analytical projects, alongside the existing portfolio scope and delivery rules.

## Excel layout

- Use a common **24 pt row height** for table rows across all worksheets. Keep rows consistent within each table and across matching tables on different sheets.
- Titles and separate embedded code blocks may retain suitable spacing. Fit wrapped headers and long labels without clipping or reducing the body font; widen columns or simplify the layout when necessary.
- Preserve the source workbook's inputs, formulas, scenario controls and financial conventions during focused changes.

## Explanation and analysis

- Add explanatory charts to the relevant statement, schedule and analytical sheets, using the Microsoft three-statement workbook as the reference for depth. Choose charts that explain trends, composition, ratios and calculation relationships; keep chart data linked to the model and selected scenario.
- Include a substantial analytical section on each meaningful analytical sheet: what the numbers show, which assumptions drive them, how the schedule connects to the statements, and material interpretation limits.
- Provide native Excel hover **Notes** for metric definitions, key inputs and important calculation cells. Explain what the metric means, its unit, where inputs come from, the actual formula and any sign, scaling, timing or allocation convention. Update Notes when a formula changes.
- Explain source actuals, calculated ratios, normalized values and forecast assumptions separately. Do not present an analyst assumption as a reported company figure.

## Raw source materials

- Include the complete original reports and raw datasets used by the project in `data/raw`. Obtain them from the issuer, regulator or the issuer's linked reporting provider.
- Preserve unmodified originals. Put annotated review copies in a separate `highlighted` directory; do not replace original pages with reconstructed tables.
- Use native PDF reports where available. If the official report is supplied only as HTML or Word, retain that original and clearly identify any PDF as a rendered copy.
- Highlight the exact reported figures or passages used by Excel. Mark used pages and add PDF bookmarks, keeping original pagination unchanged.
- Use a consistent legend: **green = income statement; blue = balance sheet; yellow = cash flow; purple = segments and note disclosures**. Put the legend in the source guide and record it in annotation metadata.
- Include a searchable source crosswalk with report year, fiscal period, document name, PDF page, source label/value, unit, Excel worksheet/cell, source URL and any scaling, aggregation or normalization.
- Use the exact comparative presentation used by the model. A FY2017 value may come from a FY2019 report; do not silently substitute an earlier presentation.
- Label calculated residuals, normalized missing rows and rounded disclosures explicitly. A missing disclosure is not proof that the issuer reported zero. Never highlight an unrelated occurrence merely because its numeric value matches.
- Record retrieval dates, filing accessions and SHA-256 hashes. Preserve the model's information cutoff and frozen source selection; date later downloads separately.

## Verification and delivery

- Verify that the highlights correspond to the selected report, period, metric and unit, and that the original and annotated PDFs have identical page counts and source text.
- Review representative marked pages visually and check the complete source crosswalk for missing or ambiguous links. Keep unresolved items visible and explain their evidence.
- Check Excel recalculation, relevant scenarios, row heights, chart bindings, Notes and preservation of existing calculations. State which native application checks were performed.
- Keep the project folder, Google Drive and GitHub synchronized with the delivered portable archive. Include source materials, the source guide, the crosswalk, reproduction utilities and verification evidence.

These are delivery requirements for future comparable projects. The earlier requirements for English deliverables, Russian collaboration, five scoping questions before a new project and the agreed project sequence remain in force.

## Mandatory corporate style references

Confirmed by Goran Brkich on 13 October 2026. These owner-supplied references take precedence over generic design defaults for future comparable projects.

- Excel reference: `Investment Banking - Quest 3.xlsx`, the original corporate workbook supplied by the owner. Retain this exact reference in the owner's files; consult it before authoring or redesigning a financial workbook.
- Word reference: `Microsoft_FY2017_10K.pdf`, the owner's supplied annual report. Derive report typography and table presentation from this document, without copying its financial content into unrelated reports.

### Excel corporate presentation

Use Arial 10 pt for body cells and bold Arial for headings. Use dark navy `#002060` with white text for main section headers. Use white backgrounds, neutral gray forecast bands and gray subtotals, thin black rules, blue editable numerical inputs and black calculation text. Preserve only functional warning or source-annotation colours. Display negative amounts in red parentheses and zero as a dash, preserving original units, precision and values. The pale yellow `#FFFF9A` from the reference is available for a selected input or sensitivity range; use it sparingly. Match charts to dark navy and neutral gray. Keep the standing uniform 24 pt table rows, native Notes, explanatory analysis and linked charts. Titles and plain code blocks may retain their necessary separate spacing.

### Word corporate presentation

Use A4 pages and the Arial typography of the supplied report. The reference's main body is 9 pt; use compact black text, bold black section headings and financial tables with white headers, alternating light gray rows and fine black horizontal rules. Main headings may be centered and displayed in capitals. Use monochrome chart presentation, white page backgrounds and restrained spacing. Do not add coloured title rules, decorative panels, gradients or additional branding colours. Preserve table and graph content, source attribution, existing section sequence and equations. Native grayscale picture effects are preferred when they preserve the original chart images unchanged.

### Focused restyling

Styling must not change inputs, formulas, scenario controls, source selections, chart bindings, native Notes or financial results. Preserve the full report text and embedded chart content. Colour legends may be updated only to describe the new style. Check all other content before and after, preserve saved numeric precision, and visually review every worksheet and report page.
