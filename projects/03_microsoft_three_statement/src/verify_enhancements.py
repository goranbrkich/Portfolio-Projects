"""Verify saved chart, Note and regression enhancements without changing Excel.

The default run uses the standard library. --run-copied-code additionally needs
NumPy and SciPy and executes a temporary copy extracted from worksheet cells.
"""
from __future__ import annotations

import argparse
import json
import math
from pathlib import Path
import posixpath
import re
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET
from zipfile import ZipFile

from workbook_io import PROJECT, NS, read_workbook, number

DRAW = "http://schemas.openxmlformats.org/drawingml/2006/main"
XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
CHART = "http://schemas.openxmlformats.org/drawingml/2006/chart"
VML = "urn:schemas-microsoft-com:vml"
EXCEL = "urn:schemas-microsoft-com:office:excel"


def close(actual, expected, label, probability=False):
    tolerance = 1e-12 + abs(expected) * 1e-7 if probability else 1e-8 * max(1, abs(expected))
    if not math.isfinite(actual) or abs(actual - expected) > tolerance:
        raise ValueError(f"{label}: saved {actual}, expected {expected}")
    return abs(actual - expected)


def sheet_parts(z):
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    rels = {r.get("Id"): r.get("Target") for r in ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))}
    return {s.get("name"): (rels[s.get("{" + NS["r"] + "}id")].lstrip("/")
            if rels[s.get("{" + NS["r"] + "}id")].startswith("/")
            else posixpath.normpath("xl/" + rels[s.get("{" + NS["r"] + "}id")]))
            for s in wb.findall("m:sheets/m:sheet", NS)}


def relationships(z, part):
    relpart = posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")
    return {r.get("Type").rsplit("/", 1)[-1]:
            (r.get("Target").lstrip("/") if r.get("Target").startswith("/")
             else posixpath.normpath(posixpath.join(posixpath.dirname(part), r.get("Target"))))
            for r in ET.fromstring(z.read(relpart))}


def baseline_check(baseline, workbook, current):
    original, _ = read_workbook(baseline)
    if list(current) != list(original) + ["Financial Modeling"]:
        raise ValueError("Original worksheet order changed")
    formula_count = populated = 0
    for sheet, cells in original.items():
        for address, old in cells.items():
            if old["value"] is None and old["formula"] is None:
                continue
            populated += 1
            new = current[sheet].get(address, {})
            if old["formula"] != new.get("formula"):
                raise ValueError(f"Original formula changed: {sheet}!{address}")
            formula_count += bool(old["formula"])
            if isinstance(old["value"], (int, float)) and not isinstance(old["value"], bool):
                if not isinstance(new.get("value"), (int, float)) or abs(old["value"] - new["value"]) > 1e-7:
                    raise ValueError(f"Original numeric value changed: {sheet}!{address}")
            elif old["value"] != new.get("value"):
                raise ValueError(f"Original text/input changed: {sheet}!{address}")
    with ZipFile(baseline) as before, ZipFile(workbook) as after:
        bp, ap = sheet_parts(before), sheet_parts(after)
        for sheet in original:
            b, a = ET.fromstring(before.read(bp[sheet])), ET.fromstring(after.read(ap[sheet]))
            for tag in ["sheetViews", "dataValidations", "sheetProtection"]:
                original_control, current_control = b.find("m:" + tag, NS), a.find("m:" + tag, NS)
                original_bytes = ET.tostring(original_control) if original_control is not None else None
                current_bytes = ET.tostring(current_control) if current_control is not None else None
                if original_bytes != current_bytes:
                    raise ValueError(f"Original {tag} changed: {sheet}")
            original_merges = {m.get("ref") for m in b.findall("m:mergeCells/m:mergeCell", NS)}
            current_merges = {m.get("ref") for m in a.findall("m:mergeCells/m:mergeCell", NS)}
            if not original_merges.issubset(current_merges):
                raise ValueError(f"Original merged cells changed: {sheet}")
    return {"original_populated_cells_preserved": populated, "original_formulas_preserved": formula_count,
            "original_sheet_order_and_controls_preserved": True}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--workbook", type=Path, default=PROJECT / "deliverables/Microsoft_Three_Statement_Model.xlsx")
    parser.add_argument("--baseline", type=Path)
    parser.add_argument("--run-copied-code", action="store_true")
    parser.add_argument("--output", type=Path, default=PROJECT / "evidence/enhancement_verification.json")
    args = parser.parse_args()
    sheets, _ = read_workbook(args.workbook)
    spec = json.loads((PROJECT / "evidence/workbook_enhancements.json").read_text())
    regression = json.loads((PROJECT / "evidence/regression/regression_results.json").read_text())
    errors = [(sn, a, c["value"]) for sn, cs in sheets.items() for a, c in cs.items() if c["type"] == "e"]
    if errors:
        raise ValueError(f"Saved formula errors: {errors}")
    notes = {}
    with ZipFile(args.workbook) as z:
        for sn, part in sheet_parts(z).items():
            rels = relationships(z, part)
            if "comments" not in rels or "vmlDrawing" not in rels:
                raise ValueError(f"Missing native hover Notes: {sn}")
            comments = ET.fromstring(z.read(rels["comments"])).findall("m:commentList/m:comment", NS)
            refs = {c.get("ref") for c in comments}
            if len(refs) != len(comments) or "C2" not in refs or any(not "".join(c.find("m:text", NS).itertext()).strip() for c in comments):
                raise ValueError(f"Invalid Note contents: {sn}")
            notes[sn] = len(comments)
            shapes = ET.fromstring(z.read(rels["vmlDrawing"])).findall("{" + VML + "}shape")
            if len(shapes) != len(comments):
                raise ValueError(f"Note popup count mismatch: {sn}")
            for shape in shapes:
                style = dict(item.strip().split(":", 1) for item in shape.get("style").split(";") if ":" in item)
                if float(style["width"].removesuffix("pt")) < 300 or float(style["height"].removesuffix("pt")) < 100 or style.get("visibility") != "hidden":
                    raise ValueError(f"Note popup geometry/visibility: {sn}")
        if notes != spec["note_counts"]:
            raise ValueError("Native Note counts differ from revision evidence")
        chart_parts = [p for p in z.namelist() if re.fullmatch(r"xl/(?:drawings/)?charts/chart\d+\.xml", p)]
        if len(chart_parts) != 37:
            raise ValueError("Expected 37 editable charts")
        for p in chart_parts:
            if not ET.fromstring(z.read(p)).findall(".//{" + CHART + "}f"):
                raise ValueError(f"Chart is not formula-linked: {p}")
        all_shapes = []
        for p in z.namelist():
            if re.fullmatch(r"xl/drawings/drawing\d+\.xml", p):
                all_shapes.extend(ET.fromstring(z.read(p)).findall(".//{" + XDR + "}sp"))
        for panel in spec["panels"]:
            match = [sp for sp in all_shapes if sp.find(".//{" + XDR + "}cNvPr").get("name") == panel["name"]]
            if not match or not any(panel["body"].replace("\n", " ").split()[0] in " ".join(sp.itertext()) for sp in match):
                raise ValueError(f"Missing expanded analytical panel: {panel['name']}")
    differences = []
    for model_name, summary, first in [("levels", 28, 12), ("first_differences", 98, 84)]:
        r = regression[model_name]
        mx = sum(o["x"] for o in r["observations"]) / r["n"]
        my = sum(o["y"] for o in r["observations"]) / r["n"]
        expected = [r["n"], r["df"], mx, my, sum((o["x"] - mx) ** 2 for o in r["observations"]),
                    r["slope"], r["intercept"], r["rss"], r["sst"], r["r_squared"], r["adjusted_r_squared"],
                    r["residual_standard_error"], r["se_ols"][1], r["se_ols"][0], r["se_hc3"][1], r["se_hc3"][0],
                    r["t_hc3"][1], r["p_hc3"][1], r["critical_t95"], *r["ci95_hc3"][1], r["durbin_watson"],
                    r["max_leverage"], r["leverage_reference"], r["t_hc3"][0], r["p_hc3"][0], *r["ci95_hc3"][0]]
        for offset, value in enumerate(expected):
            differences.append(close(number(sheets, "Financial Modeling", f"E{summary + offset}"), value,
                                     f"{model_name} summary {offset}", offset in [17, 25]))
        for index, obs in enumerate(r["observations"]):
            for column, key in [("F", "x"), ("G", "y"), ("K", "fitted"), ("L", "residual"), ("M", "leverage"), ("N", "hc3_weight")]:
                differences.append(close(number(sheets, "Financial Modeling", f"{column}{first + index}"), obs[key], f"{model_name} {index} {key}"))
    code_range = spec["python_code_range"].split("!")[1]
    start, end = map(int, re.findall(r"\d+", code_range))
    copied = "\n".join(sheets["Financial Modeling"][f"C{row}"]["value"] or "" for row in range(start, end + 1)) + "\n"
    if copied != (PROJECT / "src/regression_analysis.py").read_text():
        raise ValueError("Excel Python code differs from the supplied source")
    compile(copied, "copied_from_excel.py", "exec")
    result = {"worksheets": len(sheets), "editable_charts": len(chart_parts), "native_hover_notes": sum(notes.values()),
              "notes_per_sheet": notes, "expanded_analytical_panels": len(spec["panels"]),
              "regression_summary_comparisons": 56, "regression_observation_comparisons": len(differences) - 56,
              "maximum_regression_absolute_difference": max(differences), "saved_formula_errors": errors,
              "python_code_matches_source": True, "copied_code_executed": False}
    if args.run_copied_code:
        with tempfile.TemporaryDirectory(prefix="msft_copied_code_") as temp:
            py = Path(temp) / "regression_analysis.py"
            py.write_text(copied)
            subprocess.run([sys.executable, str(py), "--out", str(Path(temp) / "results")], check=True, capture_output=True, text=True)
            actual = json.loads((Path(temp) / "results/regression_results.json").read_text())
            if actual != regression:
                raise ValueError("Copied Python code results differ from supplied evidence")
        result["copied_code_executed"] = True
    if args.baseline:
        result.update(baseline_check(args.baseline, args.workbook, sheets))
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(result, indent=2) + "\n")
    print(json.dumps(result, indent=2))


if __name__ == "__main__":
    main()
