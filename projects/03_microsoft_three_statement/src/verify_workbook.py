"""Compare saved workbook results with independent Python calculations.

This reads cached values. Recalculate and save after edits in Excel first.
No Excel installation is required to verify the delivered saved workbook.
"""
from __future__ import annotations
import argparse
import json
from pathlib import Path
from model import forecast, validate
from workbook_io import PROJECT, read_workbook, read_excel_assumptions, number


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--workbook", type=Path, default=PROJECT / "deliverables/Microsoft_Three_Statement_Model.xlsx")
    parser.add_argument("--use-excel-inputs", action="store_true")
    parser.add_argument("--output", type=Path, default=PROJECT / "evidence/saved_workbook_verification.json")
    args = parser.parse_args()
    sheets, shape_text = read_workbook(args.workbook)
    data = json.loads((PROJECT / "data/processed/historical_financials.json").read_text())
    layout = json.loads((PROJECT / "evidence/workbook_layout.json").read_text())
    assumptions = json.loads((PROJECT / "data/processed/assumptions.json").read_text())
    case_number = number(sheets, "Assumptions", "D4")
    if case_number not in [1, 2, 3]:
        raise ValueError("Invalid case selector")
    case = ["Base", "Upside", "Downside"][int(case_number) - 1]
    if args.use_excel_inputs:
        assumptions, case = read_excel_assumptions(sheets, assumptions, layout)
    all_results = {c: forecast(data, assumptions, c) for c in ["Base", "Upside", "Downside"]}
    validation = validate(data, assumptions, all_results)
    errors = [(sheet, addr, c["value"]) for sheet, cells in sheets.items() for addr, c in cells.items() if c["type"] == "e"]
    if errors:
        raise ValueError(f"Saved formula errors: {errors}")
    audit_dependencies = [(sheet, addr, c["formula"]) for sheet, cells in sheets.items() if sheet != "Audit"
                          for addr, c in cells.items() if c["formula"] and "Audit!" in c["formula"].replace("'", "")]
    if audit_dependencies:
        raise ValueError(f"Business formulas depend on Audit: {audit_dependencies}")
    differences = []

    def compare(sheet, address, expected):
        observed = number(sheets, sheet, address)
        difference = observed - expected
        if abs(difference) > 1e-5:
            raise ValueError(f"{sheet}!{address}: saved {observed}, Python {expected}")
        differences.append({"sheet": sheet, "cell": address, "difference": difference})

    statement_maps = [("Income Statement", "IS", "income"), ("Balance Sheet", "BS", "balance"), ("Cash Flow", "CF", "cashflow")]
    for year, column in zip(range(2017, 2027), ["E", "F", "G", "H", "I", "J", "K", "L", "M", "N"]):
        for sheet, short, statement in statement_maps:
            for metric, row in layout[short].items():
                if metric in data["actuals"][str(year)][statement]:
                    if statement == "cashflow" and metric == "da_other":
                        observed = sum(number(sheets, sheet, f"{column}{r}") for r in [9, 10, 11])
                        expected = data["actuals"][str(year)][statement][metric] + data["actuals"][str(year)][statement].get("impairments", 0)
                        if abs(observed - expected) > 1e-5:
                            raise ValueError(f"{sheet}!{column}9:{column}11 D&A aggregate does not match source")
                        differences.append({"sheet": sheet, "cell": f"{column}9:{column}11", "difference": observed - expected})
                        continue
                    if statement == "cashflow" and metric == "other_investing":
                        primary = data["actuals"][str(year)][statement]
                        compare(sheet, f"{column}{row}", primary[metric] + primary.get("securities_lending_cf", 0))
                        continue
                    if statement == "cashflow" and metric == "other_financing":
                        observed = sum(number(sheets, sheet, f"{column}{r}") for r in [46, 47])
                        expected = data["actuals"][str(year)][statement][metric]
                        if abs(observed - expected) > 1e-5:
                            raise ValueError(f"{sheet}!{column}46:{column}47 financing aggregate does not match source")
                        differences.append({"sheet": sheet, "cell": f"{column}46:{column}47", "difference": observed - expected})
                        continue
                    compare(sheet, f"{column}{row}", data["actuals"][str(year)][statement][metric])
    for record, column in zip(all_results[case], ["O", "P", "Q", "R", "S"]):
        for metric in ["revenue", "cogs", "rd", "sales_marketing", "ga", "operating_income", "pretax_income", "income_tax", "net_income", "shares_diluted", "eps_diluted"]:
            compare("Income Statement", f"{column}{layout['IS'][metric]}", record[metric])
        for metric, row in layout["BS"].items():
            compare("Balance Sheet", f"{column}{row}", record["balance"][metric])
        for metric in ["net_income", "cfo", "cfi", "cff"]:
            compare("Cash Flow", f"{column}{layout['CF'][metric]}", record[metric])
        compare("Cash Flow", f"{column}{layout['CF']['capex']}", -record["cash_capex"])
        compare("Cash Flow", f"{column}{layout['CF']['closing_cash']}", record["balance"]["cash"])
    if "Scenario workflow" not in shape_text or "Cash flow signs and noncash items" not in shape_text:
        raise ValueError("Expected native calculation text boxes are absent")
    result = {"selected_case": case, "assumptions": "saved Excel inputs" if args.use_excel_inputs else "supplied JSON",
              "saved_value_comparisons": len(differences), "maximum_absolute_difference": max(abs(d["difference"]) for d in differences),
              "saved_formula_errors": errors, "audit_dependencies": audit_dependencies,
              "native_text_boxes_present": True, "python_forecast_checks": validation["forecast_checks_passed"],
              "limitations": "Reads saved values. It does not run Excel recalculation.", "comparisons": differences}
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(result, indent=2) + "\n")
    print(f"Passed {len(differences)} saved-value comparisons for {case}; {validation['forecast_checks_passed']} Python forecast checks.")


if __name__ == "__main__":
    main()
