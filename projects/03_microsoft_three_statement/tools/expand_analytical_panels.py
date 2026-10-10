"""Extend existing native DrawingML text boxes after Artifact Tool export.

Only named text-box bodies, accessibility descriptions and bounds are changed.
Charts, Notes, worksheet cells and model formulas remain untouched.
"""
import argparse
import json
import posixpath
import math
import re
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED

from lxml import etree

NS = {
    "s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}


def tag(prefix, name):
    return f"{{{NS[prefix]}}}{name}"


def expand(workbook, spec):
    entries = json.loads(Path(spec).read_text())["panels"]
    grouped = {}
    for entry in entries:
        grouped.setdefault(entry["sheet"], []).append(entry)
    with ZipFile(workbook) as archive:
        content = {name: archive.read(name) for name in archive.namelist()}
    root = etree.fromstring(content["xl/workbook.xml"])
    relations = etree.fromstring(content["xl/_rels/workbook.xml.rels"])
    targets = {r.get("Id"): r.get("Target") for r in relations}
    changed = 0
    resized_notes = 0
    for sheet in root.find("s:sheets", NS):
        name = sheet.get("name")
        target = targets[sheet.get(tag("r", "id"))]
        sheet_path = target.lstrip("/") if target.startswith("/") else "xl/" + target
        sheet_root = etree.fromstring(content[sheet_path])
        rel_path = posixpath.dirname(sheet_path) + "/_rels/" + posixpath.basename(sheet_path) + ".rels"
        rels = etree.fromstring(content[rel_path])
        def resolve(target):
            return (target.lstrip("/") if target.startswith("/") else
                    posixpath.normpath(posixpath.join(posixpath.dirname(sheet_path), target)))
        comment_rel = next((r for r in rels if r.get("Type").endswith("/comments")), None)
        vml_rel = next((r for r in rels if r.get("Type").endswith("/vmlDrawing")), None)
        if comment_rel is not None and vml_rel is not None:
            comments = etree.fromstring(content[resolve(comment_rel.get("Target"))])
            texts = {c.get("ref"): "".join(c.itertext()) for c in comments.findall("s:commentList/s:comment", NS)}
            path = resolve(vml_rel.get("Target"))
            vml = etree.fromstring(content[path])
            vns = {"v": "urn:schemas-microsoft-com:vml", "x": "urn:schemas-microsoft-com:office:excel"}
            def col_width(index):
                cols = sheet_root.find("s:cols", NS)
                if cols is not None:
                    for col in cols:
                        if int(col.get("min")) <= index + 1 <= int(col.get("max")):
                            return float(col.get("width", "8.43")) * 7 + 5
                return 64
            heights = {int(r.get("r")) - 1: float(r.get("ht")) * 4 / 3
                       for r in sheet_root.findall("s:sheetData/s:row", NS) if r.get("ht")}
            fmt = sheet_root.find("s:sheetFormatPr", NS)
            default_height = float(fmt.get("defaultRowHeight", "15")) * 4 / 3 if fmt is not None else 20
            for shape in vml.findall("v:shape", vns):
                cd = shape.find("x:ClientData", vns)
                if cd is None or cd.get("ObjectType") != "Note":
                    continue
                row, col = int(cd.findtext("x:Row", namespaces=vns)), int(cd.findtext("x:Column", namespaces=vns))
                number, letters = col + 1, ""
                while number:
                    number, rem = divmod(number - 1, 26)
                    letters = chr(65 + rem) + letters
                body = texts[f"{letters}{row + 1}"]
                lines = sum(max(1, math.ceil(len(line) / 70)) for line in body.splitlines())
                width, height = 480, max(180, min(700, lines * 18 + 32))
                style = shape.get("style", "")
                style = re.sub(r"width:[^;]+", f"width:{width * .75:g}pt", style)
                style = re.sub(r"height:[^;]+", f"height:{height * .75:g}pt", style)
                shape.set("style", style)
                anchor = cd.find("x:Anchor", vns)
                a = [int(v.strip()) for v in anchor.text.split(",")]
                end_col, x = a[0], width + a[1]
                while x >= col_width(end_col):
                    x -= col_width(end_col)
                    end_col += 1
                end_row, y = a[2], height + a[3]
                while y >= heights.get(end_row, default_height):
                    y -= heights.get(end_row, default_height)
                    end_row += 1
                anchor.text = ", ".join(map(str, a[:4] + [end_col, round(x), end_row, round(y)]))
                resized_notes += 1
            content[path] = etree.tostring(vml, encoding="UTF-8", xml_declaration=True)
        if name not in grouped:
            continue
        drawing = sheet_root.find("s:drawing", NS)
        target = next(r.get("Target") for r in rels if r.get("Id") == drawing.get(tag("r", "id")))
        draw_path = (target.lstrip("/") if target.startswith("/")
                     else posixpath.normpath(posixpath.join(posixpath.dirname(sheet_path), target)))
        drawing_root = etree.fromstring(content[draw_path])
        shapes = {sp.find("xdr:nvSpPr/xdr:cNvPr", NS).get("name"): sp
                  for sp in drawing_root.findall(".//xdr:sp", NS)}
        for entry in grouped[name]:
            shape = shapes[entry["name"]]
            props = shape.find("xdr:nvSpPr/xdr:cNvPr", NS)
            props.set("descr", entry["body"])
            shape.find("xdr:nvSpPr/xdr:cNvSpPr", NS).set("txBox", "1")
            anchor = shape.getparent()
            anchor.find("xdr:from/xdr:row", NS).text = str(entry["row"])
            extent = anchor.find("xdr:ext", NS)
            if extent is None:
                raise ValueError(f"Expected extent anchor: {name} {entry['name']}")
            for node in [extent, shape.find("xdr:spPr/a:xfrm/a:ext", NS)]:
                if node is not None:
                    node.set("cx", str(round(entry["width"] * 9525)))
                    node.set("cy", str(round(entry["height"] * 9525)))
            for old in shape.findall("xdr:txBody", NS):
                shape.remove(old)
            body = etree.SubElement(shape, tag("xdr", "txBody"))
            bp = etree.SubElement(body, tag("a", "bodyPr"), wrap="square", anchor="t",
                                 lIns="152400", rIns="152400", tIns="114300", bIns="114300")
            etree.SubElement(bp, tag("a", "noAutofit"))
            etree.SubElement(body, tag("a", "lstStyle"))
            paragraphs = [(entry["title"], True)] + [(p, False) for p in entry["body"].split("\n\n")]
            for text, bold in paragraphs:
                p = etree.SubElement(body, tag("a", "p"))
                pr = etree.SubElement(p, tag("a", "pPr"), algn="l")
                spacing = etree.SubElement(pr, tag("a", "lnSpc"))
                etree.SubElement(spacing, tag("a", "spcPct"), val="115000")
                after = etree.SubElement(pr, tag("a", "spcAft"))
                etree.SubElement(after, tag("a", "spcPts"), val="600")
                run = etree.SubElement(p, tag("a", "r"))
                rp = etree.SubElement(run, tag("a", "rPr"), lang="en-US", sz="1100" if bold else "1050", b="1" if bold else "0")
                fill = etree.SubElement(rp, tag("a", "solidFill"))
                etree.SubElement(fill, tag("a", "srgbClr"), val="17365A")
                etree.SubElement(rp, tag("a", "latin"), typeface="Arial")
                etree.SubElement(run, tag("a", "t")).text = text
                etree.SubElement(p, tag("a", "endParaRPr"), lang="en-US", sz="1050")
            changed += 1
        content[draw_path] = etree.tostring(drawing_root, encoding="UTF-8", xml_declaration=True, standalone=True)
    if changed != len(entries):
        raise ValueError("Not all requested panels were updated")
    with ZipFile(workbook, "w", ZIP_DEFLATED) as archive:
        for name, data in content.items():
            archive.writestr(name, data)
    print(f"Expanded {changed} native analytical text boxes; resized {resized_notes} Note popups")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("workbook", type=Path)
    parser.add_argument("spec", type=Path)
    args = parser.parse_args()
    expand(args.workbook, args.spec)
