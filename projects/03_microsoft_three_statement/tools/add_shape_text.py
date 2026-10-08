"""Add native editable DrawingML text to artifact-tool-created Excel shapes."""
import argparse
import json
import posixpath
import zipfile
from pathlib import Path
from lxml import etree

NS={"s":"http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r":"http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "p":"http://schemas.openxmlformats.org/package/2006/relationships",
    "xdr":"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "a":"http://schemas.openxmlformats.org/drawingml/2006/main"}
def tag(prefix, name):return f"{{{NS[prefix]}}}{name}"
def patch(workbook, notes_file):
    notes=json.loads(Path(notes_file).read_text()); grouped={}
    for note in notes:grouped.setdefault(note["sheet"],[]).append(note)
    with zipfile.ZipFile(workbook) as archive:content={name:archive.read(name) for name in archive.namelist()}
    book=etree.fromstring(content['xl/workbook.xml']);rels=etree.fromstring(content['xl/_rels/workbook.xml.rels'])
    target={r.get('Id'):r.get('Target').lstrip('/') for r in rels}
    changed=0
    for sheet in book.find('s:sheets',NS):
        name=sheet.get('name')
        if name not in grouped:continue
        dest=target[sheet.get(tag('r','id'))]
        sheet_path=dest if dest.startswith('xl/') else 'xl/'+dest
        rel_path=posixpath.dirname(sheet_path)+'/_rels/'+posixpath.basename(sheet_path)+'.rels'
        sheet_root=etree.fromstring(content[sheet_path]); drawing=sheet_root.find('s:drawing',NS)
        relation=etree.fromstring(content[rel_path]); draw_target=next(r.get('Target') for r in relation if r.get('Id')==drawing.get(tag('r','id')))
        draw_path=posixpath.normpath(posixpath.join(posixpath.dirname(sheet_path),draw_target)) if not draw_target.startswith('/') else draw_target.lstrip('/')
        root=etree.fromstring(content[draw_path]);shapes=root.findall('.//xdr:sp',NS)
        if len(shapes)!=len(grouped[name]):raise ValueError(f'Shape count mismatch on {name}')
        for shape,note in zip(shapes,grouped[name]):
            shape.find('xdr:nvSpPr/xdr:cNvPr',NS).set('name',note['title'])
            shape.find('xdr:nvSpPr/xdr:cNvPr',NS).set('descr',note['body'])
            shape.find('xdr:nvSpPr/xdr:cNvSpPr',NS).set('txBox','1')
            for old in shape.findall('xdr:txBody',NS):shape.remove(old)
            body=etree.SubElement(shape,tag('xdr','txBody'))
            bp=etree.SubElement(body,tag('a','bodyPr'),wrap='square',anchor='t',lIns='114300',rIns='114300',tIns='85725',bIns='85725')
            etree.SubElement(bp,tag('a','noAutofit'));etree.SubElement(body,tag('a','lstStyle'))
            for text,bold,size in [(note['title'],True,'1100'),(note['body'],False,'1050')]:
                p=etree.SubElement(body,tag('a','p'));pr=etree.SubElement(p,tag('a','pPr'),algn='l')
                spacing=etree.SubElement(pr,tag('a','lnSpc'));etree.SubElement(spacing,tag('a','spcPct'),val='115000')
                after=etree.SubElement(pr,tag('a','spcAft'));etree.SubElement(after,tag('a','spcPts'),val='450' if bold else '0')
                run=etree.SubElement(p,tag('a','r'));rp=etree.SubElement(run,tag('a','rPr'),lang='en-US',sz=size,b='1' if bold else '0')
                fill=etree.SubElement(rp,tag('a','solidFill'));etree.SubElement(fill,tag('a','srgbClr'),val='17365A')
                etree.SubElement(rp,tag('a','latin'),typeface='Arial');etree.SubElement(run,tag('a','t')).text=text
                etree.SubElement(p,tag('a','endParaRPr'),lang='en-US',sz=size)
            changed+=1
        content[draw_path]=etree.tostring(root,encoding='UTF-8',xml_declaration=True,standalone=True)
    calc=book.find('s:calcPr',NS)
    if calc is None:calc=etree.SubElement(book,tag('s','calcPr'))
    calc.set('calcMode','auto');calc.set('fullCalcOnLoad','1');calc.set('forceFullCalc','1')
    content['xl/workbook.xml']=etree.tostring(book,encoding='UTF-8',xml_declaration=True,standalone=True)
    with zipfile.ZipFile(workbook,'w',zipfile.ZIP_DEFLATED) as archive:
        for name,data in content.items():archive.writestr(name,data)
    print(f'Added {changed} native editable text boxes')
    return changed
if __name__=='__main__':
    parser=argparse.ArgumentParser();parser.add_argument('workbook',type=Path);parser.add_argument('notes',type=Path);args=parser.parse_args();patch(args.workbook,args.notes)
