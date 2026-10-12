"""Create the source-review guide from the verified PDF crosswalk."""
from collections import Counter
import json
from pathlib import Path
from xml.sax.saxutils import escape

from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, KeepTogether

PROJECT=Path(__file__).resolve().parents[1]


def page_ranges(values):
    values=sorted(set(values));parts=[]
    for value in values:
        if parts and value==parts[-1][-1]+1:parts[-1].append(value)
        else:parts.append([value])
    return ', '.join(str(p[0]) if len(p)==1 else f'{p[0]}-{p[-1]}' for p in parts)


def main():
    plan=json.loads((PROJECT/'evidence/raw_materials_mapping.json').read_text())
    verify=json.loads((PROJECT/'evidence/raw_materials_verification.json').read_text())
    manifest=json.loads((PROJECT/'data/raw/raw_materials_manifest.json').read_text())
    records=plan['source_records'];out=PROJECT/'docs/Microsoft_Source_Guide.pdf'
    styles=getSampleStyleSheet();styles.add(ParagraphStyle(name='Body',fontName='Helvetica',fontSize=10,leading=14,spaceAfter=8,textColor=colors.HexColor('#263649')))
    styles.add(ParagraphStyle(name='SmallBody',parent=styles['Body'],fontSize=8.4,leading=11,spaceAfter=3))
    styles.add(ParagraphStyle(name='TableHead',parent=styles['SmallBody'],textColor=colors.white,fontName='Helvetica-Bold'))
    styles['Title'].fontName='Helvetica-Bold';styles['Title'].fontSize=20;styles['Title'].leading=24;styles['Title'].textColor=colors.HexColor('#17365A')
    styles['Heading2'].fontName='Helvetica-Bold';styles['Heading2'].fontSize=13;styles['Heading2'].leading=17;styles['Heading2'].textColor=colors.HexColor('#17365A');styles['Heading2'].spaceBefore=12
    def para(t,style='Body'):return Paragraph(t,styles[style])
    def table(rows,widths,header=True):
        t=Table([[para(str(c),'TableHead' if header and i==0 else 'SmallBody') for c in row] for i,row in enumerate(rows)],colWidths=widths,repeatRows=1 if header else 0,hAlign='LEFT')
        rules=[('VALIGN',(0,0),(-1,-1),'TOP'),('LEFTPADDING',(0,0),(-1,-1),7),('RIGHTPADDING',(0,0),(-1,-1),7),('TOPPADDING',(0,0),(-1,-1),4),('BOTTOMPADDING',(0,0),(-1,-1),4),('LINEBELOW',(0,0),(-1,0),.5,colors.HexColor('#17365A'))]
        if header:rules.extend([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#17365A')),('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white,colors.HexColor('#F2F5F8')])])
        t.setStyle(TableStyle(rules));return t
    story=[para('Microsoft<br/>Source Review Guide','Title'),para('Portfolio project 03. Source package added 12 October 2026. Model cutoff remains 8 October 2026.','SmallBody'),Spacer(1,5*mm)]
    story.append(para('How to check a workbook number','Heading2'))
    story.append(para('Open <b>source_crosswalk.csv</b>, find the Excel worksheet and cell, then open the corresponding highlighted filing at the listed <b>PDF page</b>. Hover or click the colored annotation to see the source ID, fiscal period, metric, source value and Excel references. The purple or other colored margin stripe identifies a page used by the workbook. PDF bookmarks jump directly to used pages.'))
    story.append(para('The package contains ten complete native Form 10-K PDFs, their annotated copies, and thirteen original-format HTML/JSON snapshots. Originals are in <b>data/raw/reports/originals</b>; review copies are in <b>data/raw/reports/highlighted</b>. Original pagination and source text are retained.'))
    legend=[['Color','Meaning'],['Green','Income-statement inputs: revenue, costs, earnings and share counts'],['Blue','Balance-sheet inputs, including FY2016 opening-cash context'],['Yellow','Cash-flow inputs and financing/investing components'],['Purple','Segments, note disclosures and contractual payment schedules']]
    lt=table(legend,[29*mm,143*mm])
    for row,color in enumerate([(0.55,0.92,0.64),(0.59,0.80,0.98),(1,.9,.55),(.84,.72,.97)],1):lt.setStyle(TableStyle([('BACKGROUND',(0,row),(0,row),colors.Color(*color))]))
    story.extend([para('Color legend','Heading2'),lt])
    story.append(para('Reported values and model conventions','Heading2'))
    story.append(para('The fiscal period and filing year can differ. The model uses the latest available comparative presentation at its cutoff. For example, FY2017 revenue is taken from the FY2019 filing, while the FY2017 balance sheet is from FY2018. FY2016 is opening-balance context only, not an eleventh historical year.'))
    story.append(para('Some note figures are disclosed in billions and stored in Excel in millions. Certain contra balances and expenses are shown as negative deductions in the report and stored as positive magnitudes in the schedule. The crosswalk records both values and labels these conversions.'))
    story.append(para('A normalized zero for an omitted line is not an explicit company disclosure. Calculated cash-flow residuals and aggregates are also identified separately. These records point to the relevant source table and are not highlighted as an invented reported number. Forecast assumptions are analyst inputs; no source annotation presents them as reported actuals.'))
    story.append(PageBreak())
    story.append(para('Report index and marked pages','Title'))
    index=[['Form 10-K','PDF pages','Used PDF pages','Source records']]
    for report in verify['reports']:
        y=report['report_year'];count=sum(r['report_year']==y for r in records)
        index.append([f'FY{y}',plan['report_pages'][str(y)],page_ranges(report['used_pdf_pages']),count])
    story.append(table(index,[27*mm,23*mm,98*mm,24*mm]))
    story.append(Spacer(1,4*mm))
    story.append(para('PDF page means the one-based page shown by the PDF viewer. It is kept unchanged between each original and its highlighted copy. Some filings include exhibits after the financial statements.','SmallBody'))
    story.append(para('Verification','Heading2'))
    n=sum(r['value_and_label_annotations'] for r in verify['reports']);used=sum(len(r['used_pdf_pages']) for r in verify['reports'])
    story.append(para(f'<b>{len(records):,}</b> source records link workbook inputs to evidence. The copies contain <b>{n:,}</b> value/label highlights across <b>{used}</b> used report pages. Source text and page counts match the originals. Excel and Word deliverables are unchanged.'))
    statuses=Counter(r['status'] for r in records)
    story.append(table([['Evidence convention','Records']]+[[k.replace('_',' '),v] for k,v in sorted(statuses.items())],[140*mm,32*mm]))
    story.append(para('Download timing and reproducibility','Heading2'))
    story.append(para('New downloads are dated separately from the original model sources. The initial source_manifest.json and the frozen selected_sec_facts.json remain intact. raw_materials_manifest.json records source URLs, filing accessions, retrieval times and SHA-256 hashes.'))
    changed=[r['path'].split('/')[-1] for r in manifest['sources'] if r.get('same_bytes_as_initial_model_source') is False]
    if changed:story.append(para('The freshly downloaded '+', '.join(changed)+' differs in bytes from the initial capture. The frozen extracted statement tables and selected XBRL facts retain the model\'s source selection; the PDF links were verified against its actual input values.','SmallBody'))
    story.append(PageBreak())
    story.append(para('Examples and source links','Title'))
    samples=[]
    for cell in ['Historical Inputs!E9','Historical Inputs!N9','Historical Inputs!N80','Filing Notes!N25','Filing Notes!O37','Filing Notes!N31']:
        r=next(r for r in records if cell in r['cells']);samples.append(r)
    for r in samples:
        story.append(para(f'{escape(r["cells"][0])}: {escape(r["metric"].replace("_"," "))}','Heading2'))
        story.append(para(f'FY{r["fiscal_year"]}; FY{r["report_year"]} filing; PDF page {r["page"]}. Source value {r["source_value"]:,}; workbook value {r["workbook_value"]:,} {r["unit"]}. Convention: {r["status"].replace("_"," ")}.','SmallBody'))
    story.append(para('Official filing downloads','Heading2'))
    for record in manifest['sources']:
        if record['kind']=='pdf':
            url=record['url'];story.append(para(f'FY{record["report_year"]}: <link href="{url}" color="#235C8A">{record["accession"]}</link>','SmallBody'))
    story.append(para('Standing requirements for future projects','Heading2'))
    story.append(para('The shared PROJECT_STANDARDS.md and root AGENTS.md require 24 pt Excel table rows, explanatory charts and expanded analysis, native hover Notes, complete raw materials, separate highlighted PDFs and a source crosswalk for future similar projects.'))
    def footer(canvas,doc):
        canvas.setFont('Helvetica',8);canvas.setFillColor(colors.HexColor('#687888'));canvas.drawString(19*mm,12*mm,'Goran Brkich | Microsoft source review');canvas.drawRightString(191*mm,12*mm,str(doc.page))
    SimpleDocTemplate(str(out),pagesize=(210*mm,297*mm),leftMargin=19*mm,rightMargin=19*mm,topMargin=18*mm,bottomMargin=20*mm,title='Microsoft Source Review Guide',author='Goran Brkich').build(story,onFirstPage=footer,onLaterPages=footer)
    print(out)


if __name__=='__main__':main()
