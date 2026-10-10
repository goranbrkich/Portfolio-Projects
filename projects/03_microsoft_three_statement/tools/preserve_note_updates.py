"""Consolidate duplicate native Notes emitted by the authoring API.

The API authors the replacement text but adds another Note at the same cell.
Keep the existing native annotation, copy in the authored replacement body,
and retain one VML popup per cell. No worksheet cells or formulas are edited.
"""
import argparse
from copy import deepcopy
import json
from pathlib import Path
import posixpath
from zipfile import ZipFile, ZIP_DEFLATED

from lxml import etree

MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
VML = 'urn:schemas-microsoft-com:vml'
EXCEL = 'urn:schemas-microsoft-com:office:excel'
NS = {'m': MAIN, 'v': VML, 'x': EXCEL}


def text(comment):
    return ''.join(comment.find('m:text', NS).itertext())


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('workbook', type=Path)
    parser.add_argument('plan', type=Path)
    args = parser.parse_args()
    edits = json.loads(args.plan.read_text())['observationEdits']
    targets = {(e['sheet'], e['cell']): e for e in edits}
    with ZipFile(args.workbook) as z:
        content = {n: z.read(n) for n in z.namelist()}
    wb = etree.fromstring(content['xl/workbook.xml'])
    wb_rels = {r.get('Id'): r.get('Target') for r in etree.fromstring(content['xl/_rels/workbook.xml.rels'])}
    removed = popups = 0
    for sheet in wb.findall('m:sheets/m:sheet', NS):
        sn = sheet.get('name')
        target = wb_rels[sheet.get('{' + REL + '}id')]
        part = target.lstrip('/') if target.startswith('/') else posixpath.normpath('xl/' + target)
        rel_path = posixpath.join(posixpath.dirname(part), '_rels', posixpath.basename(part) + '.rels')
        rels = {r.get('Type').rsplit('/', 1)[-1]: (r.get('Target').lstrip('/') if r.get('Target').startswith('/') else posixpath.normpath(posixpath.join(posixpath.dirname(part), r.get('Target')))) for r in etree.fromstring(content[rel_path])}
        root = etree.fromstring(content[rels['comments']])
        comments = root.find('m:commentList', NS)
        grouped = {}
        for c in comments:
            grouped.setdefault(c.get('ref'), []).append(c)
        for ref, group in grouped.items():
            if len(group) == 1:
                continue
            edit = targets.get((sn, ref))
            if edit is None or len(group) != 2:
                raise ValueError(f'Unexpected duplicate Note: {sn}!{ref}')
            original = next((c for c in group if text(c) == edit['oldNote']), None)
            replacement = next((c for c in group if text(c) == edit['newNote']), None)
            if original is None or replacement is None:
                raise ValueError(f'Missing original/replacement Note: {sn}!{ref}')
            original.replace(original.find('m:text', NS), deepcopy(replacement.find('m:text', NS)))
            for c in group:
                if c is not original:
                    comments.remove(c)
                    removed += 1
        content[rels['comments']] = etree.tostring(root, encoding='UTF-8', xml_declaration=True)
        vml = etree.fromstring(content[rels['vmlDrawing']])
        seen = set()
        for shape in vml.findall('v:shape', NS):
            cd = shape.find('x:ClientData', NS)
            if cd is None or cd.get('ObjectType') != 'Note':
                continue
            key = (int(cd.findtext('x:Row', namespaces=NS)), int(cd.findtext('x:Column', namespaces=NS)))
            if key in seen:
                vml.remove(shape)
                popups += 1
            else:
                seen.add(key)
        content[rels['vmlDrawing']] = etree.tostring(vml, encoding='UTF-8', xml_declaration=True)
    if removed != len(targets) or popups != len(targets):
        raise ValueError(f'Expected {len(targets)} Note updates, found {removed}/{popups}')
    with ZipFile(args.workbook, 'w', ZIP_DEFLATED) as z:
        for name, data in content.items():
            z.writestr(name, data)
    print(f'Updated {removed} existing Notes; consolidated {popups} popup duplicates.')


if __name__ == '__main__':
    main()
