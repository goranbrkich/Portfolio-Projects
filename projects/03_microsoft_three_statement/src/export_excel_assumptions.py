"""Export saved Excel inputs to JSON without changing the workbook."""
from __future__ import annotations
import argparse
import json
from pathlib import Path
from workbook_io import PROJECT, read_workbook, read_excel_assumptions


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--workbook", type=Path, default=PROJECT / "deliverables/Microsoft_Three_Statement_Model.xlsx")
    parser.add_argument("--output", type=Path, default=PROJECT / "evidence/excel_assumptions.json")
    args = parser.parse_args()
    sheets, _ = read_workbook(args.workbook)
    template = json.loads((PROJECT / "data/processed/assumptions.json").read_text())
    layout = json.loads((PROJECT / "evidence/workbook_layout.json").read_text())
    assumptions, active_case = read_excel_assumptions(sheets, template, layout)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(assumptions, indent=2) + "\n")
    print(f"Exported all scenario inputs to {args.output}. Selected Excel case: {active_case}.")


if __name__ == "__main__":
    main()
