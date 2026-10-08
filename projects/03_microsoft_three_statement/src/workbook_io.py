"""Read saved OOXML values without modifying the Excel workbook."""
from __future__ import annotations

import copy
import json
from pathlib import Path
import posixpath
import xml.etree.ElementTree as ET
import zipfile

PROJECT = Path(__file__).resolve().parents[1]
NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
PKG = {"p": "http://schemas.openxmlformats.org/package/2006/relationships"}


def read_workbook(path):
    with zipfile.ZipFile(path) as z:
        if z.testzip() is not None:
            raise ValueError("Corrupted workbook archive")
        wb = ET.fromstring(z.read("xl/workbook.xml"))
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        targets = {r.attrib["Id"]: r.attrib["Target"] for r in rels}
        strings = []
        if "xl/sharedStrings.xml" in z.namelist():
            strings = ["".join(si.itertext()) for si in ET.fromstring(z.read("xl/sharedStrings.xml"))]
        sheets = {}
        for sheet in wb.findall("m:sheets/m:sheet", NS):
            target = targets[sheet.attrib["{" + NS["r"] + "}id"]]
            entry = target.lstrip("/") if target.startswith("/") else posixpath.normpath("xl/" + target)
            root = ET.fromstring(z.read(entry))
            cells = {}
            for cell in root.findall(".//m:sheetData/m:row/m:c", NS):
                raw = cell.findtext("m:v", default=None, namespaces=NS)
                kind = cell.attrib.get("t", "n")
                if kind == "s":
                    value = strings[int(raw)] if raw is not None else None
                elif kind == "inlineStr":
                    value = "".join(cell.find("m:is", NS).itertext())
                elif kind == "b":
                    value = raw == "1"
                elif kind in {"str", "e"}:
                    value = raw
                else:
                    value = float(raw) if raw is not None else None
                cells[cell.attrib["r"]] = {"value": value, "type": kind,
                                            "formula": cell.findtext("m:f", default=None, namespaces=NS)}
            sheets[sheet.attrib["name"]] = cells
        shape_text = []
        for entry in z.namelist():
            if entry.startswith("xl/drawings/drawing") and entry.endswith(".xml"):
                root = ET.fromstring(z.read(entry))
                shape_text.extend(t.text for t in root.findall(".//{http://schemas.openxmlformats.org/drawingml/2006/main}t") if t.text)
        return sheets, shape_text


def number(sheets, sheet, address):
    value = sheets[sheet].get(address, {}).get("value")
    if not isinstance(value, (float, int)) or isinstance(value, bool):
        raise ValueError(f"Missing numeric saved value: {sheet}!{address}")
    return value


def read_excel_assumptions(sheets, template, layout):
    assumptions = copy.deepcopy(template)
    case_number = number(sheets, "Assumptions", "D4")
    if case_number not in (1, 2, 3):
        raise ValueError("Assumptions D4 must be 1, 2 or 3")
    for driver, active_row in layout["driver_rows"].items():
        for offset, case in enumerate(["Base", "Upside", "Downside"], start=1):
            assumptions["cases"][case][driver] = [
                number(sheets, "Assumptions", f"{column}{active_row + offset}")
                for column in ["O", "P", "Q", "R", "S"]]
    for driver in ["ppe", "intangibles", "operating_leases"]:
        row = layout["allocation_rows"][driver]
        assumptions["allocation"][driver] = [number(sheets, "Assumptions", f"{col}{row}") for col in ["E", "F", "G", "H"]]
    for key, short in [("fy2032_debt_maturity_assumption", "debt2032"),
                       ("fy2032_legacy_operating_lease_payment_assumption", "op2032"),
                       ("fy2032_legacy_finance_lease_payment_assumption", "fin2032")]:
        assumptions[key] = number(sheets, "Assumptions", f"O{layout['common_rows'][short]}")
    return assumptions, ["Base", "Upside", "Downside"][int(case_number) - 1]
