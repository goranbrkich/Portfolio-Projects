"""Verify the locale-safe observations and common table row height.

Read-only OOXML checks use the standard library. --baseline additionally
compares the preceding workbook, allowing only the specified observation edits.
--native-workbook compares a disposable copy recalculated by an office engine.
"""
from __future__ import annotations

import argparse
import json
import math
from pathlib import Path
from zipfile import ZipFile
import xml.etree.ElementTree as ET

from workbook_io import PROJECT, NS, read_workbook
from verify_enhancements import sheet_parts, relationships


def same_value(actual, expected):
    if isinstance(expected, (int, float)) and not isinstance(expected, bool):
        return isinstance(actual, (int, float)) and not isinstance(actual, bool) and math.isclose(actual, expected, rel_tol=1e-12, abs_tol=1e-7)
    return actual == expected


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--workbook', type=Path, default=PROJECT / 'deliverables/Microsoft_Three_Statement_Model.xlsx')
    parser.add_argument('--baseline', type=Path)
    parser.add_argument('--native-workbook', type=Path)
    parser.add_argument('--output', type=Path, default=PROJECT / 'evidence/excel_compatibility_verification.json')
    args = parser.parse_args()
    plan = json.loads((PROJECT / 'evidence/excel_compatibility_plan.json').read_text())
    sheets, _ = read_workbook(args.workbook)
    targets = {(e['sheet'], e['cell']): e for e in plan['observationEdits']}
    errors = [(sn, a, c['value']) for sn, cells in sheets.items() for a, c in cells.items() if c['type'] == 'e']
    if errors:
        raise ValueError(f'Saved formula errors: {errors[:5]}')
    for (sn, a), e in targets.items():
        cell = sheets[sn][a]
        if '=' + cell['formula'] != e['newFormula'] or 'TEXT(' in cell['formula'] or not isinstance(cell['value'], str):
            raise ValueError(f'Observation formula: {sn}!{a}')
        if args.baseline and cell['value'] != e['oldValue']:
            raise ValueError(f'Observation meaning changed: {sn}!{a}: {cell["value"]}')
    rows_checked = notes_checked = 0
    with ZipFile(args.workbook) as z:
        parts = sheet_parts(z)
        for row_plan in plan['rowPlans']:
            sn = row_plan['sheet']
            root = ET.fromstring(z.read(parts[sn]))
            heights = {int(row.get('r')): float(row.get('ht', '15')) for row in root.findall('m:sheetData/m:row', NS)}
            for row in range(row_plan['firstRow'], row_plan['lastRow'] + 1):
                if heights.get(row) != plan['tableRowHeightPoints']:
                    raise ValueError(f'Inconsistent height: {sn}!row {row}: {heights.get(row)}')
                rows_checked += 1
            rels = relationships(z, parts[sn])
            comments = ET.fromstring(z.read(rels['comments'])).findall('m:commentList/m:comment', NS)
            note_refs = [c.get('ref') for c in comments]
            if len(note_refs) != len(set(note_refs)):
                raise ValueError(f'Duplicate hover Notes: {sn}')
            for comment in comments:
                edit = targets.get((sn, comment.get('ref')))
                if edit is not None:
                    if ''.join(comment.find('m:text', NS).itertext()) != edit['newNote']:
                        raise ValueError(f'Outdated observation Note: {sn}!{comment.get("ref")}')
                    notes_checked += 1
    if notes_checked != len(targets):
        raise ValueError('Missing observation Notes')
    result = {'observation_formulas': len(targets), 'format_calls_replaced': plan['textCallsReplaced'],
              'table_row_height_points': plan['tableRowHeightPoints'], 'sheets_with_common_height': len(plan['rowPlans']),
              'normalized_rows_checked': rows_checked, 'observation_notes_updated': notes_checked,
              'saved_formula_errors': errors, 'baseline_compared': False, 'native_engine_compared': False,
              'native_engine': None, 'microsoft_desktop_excel_tested': False}
    if args.baseline:
        baseline, _ = read_workbook(args.baseline)
        if list(baseline) != list(sheets):
            raise ValueError('Worksheet order changed')
        preserved = formulas = 0
        for sn, cells in baseline.items():
            for a, old in cells.items():
                if old['value'] is None and old['formula'] is None:
                    continue
                new = sheets[sn].get(a)
                if new is None or not same_value(new['value'], old['value']):
                    raise ValueError(f'Value changed outside scope: {sn}!{a}')
                if (sn, a) not in targets and new['formula'] != old['formula']:
                    raise ValueError(f'Formula changed outside scope: {sn}!{a}')
                if (sn, a) not in targets:
                    preserved += 1
                    formulas += bool(old['formula'])
        with ZipFile(args.baseline) as before, ZipFile(args.workbook) as after:
            bp, ap = sheet_parts(before), sheet_parts(after)
            for sn in sheets:
                b, a = ET.fromstring(before.read(bp[sn])), ET.fromstring(after.read(ap[sn]))
                for tag in ['sheetViews', 'dataValidations', 'sheetProtection', 'mergeCells', 'conditionalFormatting', 'cols']:
                    bv = [ET.tostring(n) for n in b.findall('m:' + tag, NS)]
                    av = [ET.tostring(n) for n in a.findall('m:' + tag, NS)]
                    if bv != av:
                        raise ValueError(f'Control/style changed outside scope: {sn} {tag}')
        result.update(baseline_compared=True, untouched_populated_cells_preserved=preserved,
                      untouched_formulas_preserved=formulas)
    if args.native_workbook:
        native, _ = read_workbook(args.native_workbook)
        native_errors = [(sn, a, c['value']) for sn, cells in native.items() for a, c in cells.items() if c['type'] == 'e']
        if native_errors:
            raise ValueError(f'Native recalculation errors: {native_errors[:5]}')
        for (sn, a), e in targets.items():
            if not same_value(native[sn][a]['value'], sheets[sn][a]['value']):
                raise ValueError(f'Native observation mismatch: {sn}!{a}')
        result.update(native_engine_compared=True, native_engine='LibreOffice Calc', native_saved_formula_errors=native_errors)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(result, indent=2) + '\n')
    print(json.dumps(result, indent=2))


if __name__ == '__main__':
    main()
