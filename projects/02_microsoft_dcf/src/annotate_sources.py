"""Create marked review copies with DCF cell references; originals remain intact."""
from __future__ import annotations
import csv
from collections import defaultdict
import hashlib
import json
from pathlib import Path
import re
import sys

import fitz
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak

ROOT = Path(__file__).resolve().parents[1]
COLORS = {'income': (0.35, .86, .50), 'balance': (.32, .68, .96),
          'cashflow': (1, .82, .20), 'notes': (.70, .48, .92), 'segments': (.70, .48, .92)}

def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()

def main():
    prior = Path(sys.argv[1])
    rows = list(csv.DictReader((ROOT / 'docs/source_crosswalk.csv').open(encoding='utf-8-sig')))
    # Preserve explicit source units for the two additional DCF observations.
    for r in rows:
        if r['source_id'] == 'DCF-SHARES':
            r['reported_pdf_value'], r['reported_pdf_unit'] = '7425545491', 'shares'
        if r['source_id'] == 'DCF-DEBT-FV':
            r['reported_pdf_value'], r['reported_pdf_unit'] = '36.5', 'USD billions'
    with (ROOT / 'docs/source_crosswalk.csv').open('w', encoding='utf-8-sig', newline='') as out:
        writer = csv.DictWriter(out, fieldnames=list(rows[0])); writer.writeheader(); writer.writerows(rows)
    selected = {r['source_id']: r for r in rows if r['status'].startswith('reported')}
    by_report = defaultdict(list)
    for r in selected.values():
        by_report[r['report_year']].append(r)
    dest = ROOT / 'data/raw/reports/highlighted'; dest.mkdir(parents=True, exist_ok=True)
    results = []; found = set()
    for year in range(2017, 2027):
        filename = f'Microsoft_FY{year}_10K.pdf'
        original = ROOT / 'data/raw/reports/originals' / filename
        original_hash = sha(original)
        doc = fitz.open(original)
        old = fitz.open(prior / 'data/raw/reports/highlighted' / filename.replace('.pdf', '_highlighted.pdf'))
        count = 0; used = defaultdict(set)
        for i, page in enumerate(old):
            for annot in page.annots() or []:
                body = annot.info.get('content', '')
                match = re.match(r'(SRC-\d+):', body)
                if annot.type[0] != 8 or not match or match.group(1) not in selected:
                    continue
                source = selected[match.group(1)]
                if int(source['report_year']) != year or int(source['pdf_page']) != i + 1:
                    raise AssertionError('Source ID points to another report or page')
                quads = [fitz.Quad(annot.vertices[j:j+4]) for j in range(0, len(annot.vertices), 4)]
                target_page = doc[i]
                new = target_page.add_highlight_annot(quads)
                new.set_colors(stroke=COLORS[source['category']]); new.set_opacity(.32)
                new.set_info(title='Goran Brkich Portfolio', subject=source['category'],
                             content=f"{source['source_id']}: FY{source['fiscal_year']} {source['source_label']}; "
                                     f"source {source['reported_pdf_value'] or source['source_value']} {source['reported_pdf_unit']}; "
                                     f"Excel {source['workbook_value']} {source['unit']}; {source['excel_cells']}; "
                                     f"{source['conversion']}; {source['status']}")
                new.update(); count += 1; used[i+1].add(source['source_id']); found.add(source['source_id'])
        if year == 2026:
            for identifier, token in [('DCF-SHARES','7,425,545,491'), ('DCF-DEBT-FV','36.5')]:
                source = selected[identifier]; pno = int(source['pdf_page']) - 1; page = doc[pno]
                candidates = page.search_for(token)
                if identifier == 'DCF-DEBT-FV':
                    candidates = [r for r in candidates if 'fair value' in page.get_textbox(fitz.Rect(0,r.y0-28,page.rect.width,r.y1+28)).lower()]
                if len(candidates) != 1:
                    raise AssertionError(f'Expected one contextual source for {identifier}: {candidates}')
                a = page.add_highlight_annot(candidates[0]); a.set_colors(stroke=COLORS['notes']); a.set_opacity(.32)
                a.set_info(title='Goran Brkich Portfolio', subject='Additional DCF observation',
                           content=f"{identifier}: {source['source_label']}; {source['source_value']} {source['reported_pdf_unit']}; "
                                   f"{source['conversion']}; Excel {source['workbook_value']} {source['unit']}; {source['excel_cells']}")
                a.update(); count += 1; used[pno+1].add(identifier); found.add(identifier)
        for page_number, identifiers in used.items():
            target_page = doc[page_number-1]
            stripe = target_page.add_rect_annot(fitz.Rect(3,26,7,73)); stripe.set_colors(stroke=COLORS['notes'],fill=COLORS['notes']); stripe.set_opacity(.60)
            stripe.set_info(title='Page used in Microsoft DCF',content='; '.join(sorted(identifiers))); stripe.update()
        doc.set_toc([[1,f'Used PDF page {page}: {len(ids)} input records',page] for page,ids in sorted(used.items())])
        output = dest / filename.replace('.pdf','_highlighted.pdf')
        doc.save(output, garbage=2, deflate=True, use_objstms=1); doc.close(); old.close()
        with fitz.open(original) as a, fitz.open(output) as b:
            assert len(a)==len(b)
            assert all(x.get_text()==y.get_text() for x,y in zip(a,b))
            assert sha(original)==original_hash
            result = {'report_year':year,'original':str(original.relative_to(ROOT)),'highlighted':str(output.relative_to(ROOT)),
                      'pages':len(a),'used_pdf_pages':sorted(used),'source_record_count':len(set().union(*used.values())) if used else 0,
                      'highlight_annotations':count,'original_sha256':original_hash,'highlighted_sha256':sha(output),
                      'original_unchanged':True,'native_text_equal':True,'bytes':output.stat().st_size}
            results.append(result)
    missing = sorted(set(selected)-found)
    if missing:
        raise AssertionError(f'No annotation for reported source records: {missing}')
    verification={'reports':results,'reported_source_records':len(selected),'annotated_source_records':len(found),
                  'source_crosswalk_records':len(rows),'unavailable_disclosures':sum(r['status']=='unavailable_disclosure' for r in rows)}
    (ROOT/'evidence/raw_materials_checks.json').write_text(json.dumps(verification,indent=2))
    source_guide(verification)
    market_extracts()
    manifest=[]
    for p in sorted((ROOT/'data/raw').rglob('*')):
        if p.is_file() and p.name!='source_manifest.json':
            manifest.append({'path':str(p.relative_to(ROOT)),'sha256':sha(p),'bytes':p.stat().st_size})
    (ROOT/'data/raw/source_manifest.json').write_text(json.dumps({'information_cutoff':'2026-10-08','files':manifest},indent=2))
    print(json.dumps({'reports':len(results),'reported_sources':len(found),'annotations':sum(x['highlight_annotations'] for x in results)}))

styles=getSampleStyleSheet()
styles.add(ParagraphStyle(name='Copy',fontName='Helvetica',fontSize=10,leading=14,spaceAfter=9))
styles.add(ParagraphStyle(name='BlackTitle',fontName='Helvetica-Bold',fontSize=20,leading=24,spaceAfter=16,textColor=colors.black))
styles.add(ParagraphStyle(name='BlackHead',fontName='Helvetica-Bold',fontSize=13,leading=17,spaceAfter=8,textColor=colors.black))
def para(text, style='Copy'):
    return Paragraph(text, styles[style])
def doc_pdf(name, story):
    def footer(c,d):
        c.setFont('Helvetica',8);c.drawString(48,28,'Goran Brkich | Microsoft DCF | Information cutoff 8 October 2026');c.drawRightString(564,28,str(d.page))
    SimpleDocTemplate(str(ROOT/'deliverables'/name),pagesize=letter,rightMargin=48,leftMargin=48,topMargin=45,bottomMargin=45).build(story,onFirstPage=footer,onLaterPages=footer)
def source_guide(v):
    story=[para('Microsoft source evidence guide','BlackTitle'),para('Portfolio project 02 | 8 October 2026 cutoff'),
           para('The full original reports are in <b>data/raw/reports/originals</b>. Review copies in <b>data/raw/reports/highlighted</b> preserve all pages and the original native text. Colored annotations identify the exact source values and their Excel destinations.'),
           para('Find an Excel cell in <b>docs/source_crosswalk.csv</b> or Source Catalog, then open the marked PDF at its one-based PDF page. Hover or click an annotation for source ID, fiscal year, units, conversion and Excel reference. Bookmarks and margin stripes identify used pages.'),
           para('Highlight colours','BlackHead')]
    table=Table([['Category','Colour','Examples'],['Income','Green','Revenue and operating profit'],['Balance','Blue','Cash, NWC assets and liabilities'],['Cash flow','Yellow','CFO, cash capex and SBC'],['Notes and segments','Purple','Leases, depreciation, shares and debt fair value']],colWidths=[115,90,311])
    table.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#EAF0F7')),('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),('FONTSIZE',(0,0),(-1,-1),9),('BOTTOMPADDING',(0,0),(-1,-1),8),('TOPPADDING',(0,0),(-1,-1),8),('LINEBELOW',(0,0),(-1,0),.6,colors.grey)]))
    story+=[table,Spacer(1,14),para(f"The crosswalk contains <b>{v['source_crosswalk_records']}</b> records. All <b>{v['reported_source_records']}</b> reported records have source annotations. <b>{v['unavailable_disclosures']}</b> missing prior note disclosures remain n.a. Seven normalized absence records identify omitted lines; they are not highlighted as explicit reported numbers."),
            para('Historical period selection','BlackHead'),para('FY2017 to FY2026 financial history uses the selected later comparative filings inherited from project 03. This preserves segment recasts and statement presentation. Report year can therefore differ from the fiscal year of an input. The FY2017 PDF is retained as an archive original but has no selected source input in this DCF crosswalk.'),PageBreak(),para('Report pages used in the workbook','BlackTitle')]
    report_table=[['Report','PDF pages used','Input records']]+[[f"FY{x['report_year']}",', '.join(map(str,x['used_pdf_pages'])) or 'Archive only',str(x['source_record_count'])] for x in v['reports']]
    t=Table(report_table,colWidths=[80,350,86]);t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#EAF0F7')),('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),('FONTSIZE',(0,0),(-1,-1),9),('TOPPADDING',(0,0),(-1,-1),7),('BOTTOMPADDING',(0,0),(-1,-1),7)]));story+=[t,Spacer(1,14),para('Market and method evidence','BlackHead'),para('Original quote responses, Treasury CSV and HTML, Damodaran ERP and rating HTML, calculator workbooks and HC3 method documentation are preserved in data/raw/market. Market_Source_Extracts.pdf is a readable review extract generated from those originals, clearly labelled as an extract.'),para('The valuation selects the 8 October regular close and same-date Treasury yield. The explicitly dated 1 October ERP is used. The monthly ERP archive has an inconsistent last date serial and is retained as supporting evidence rather than used to set the valuation period.'),para('Verification','BlackHead'),para('The evidence manifest records file sizes and SHA-256 hashes. The annotation check verifies all native PDF text, page counts and original hashes remain unchanged. Original PDFs and raw HTML/JSON/CSV are separate from marked or generated review copies.')]
    doc_pdf('Source_Evidence_Guide.pdf',story)
def market_extracts():
    market=json.loads((ROOT/'data/processed/market_inputs.json').read_text())
    story=[para('Market source review extracts','BlackTitle'),para('Generated review extract | Original HTML JSON CSV and XLSX remain in data/raw/market'),para('Microsoft historic price observation','BlackHead'),para('Issuer-linked Microsoft stock lookup<br/>Requested date 8 October 2026<br/><b>Regular closing price USD 522.61</b><br/>Reported volume 20,218,987 shares'),para('Original response microsoft_stock_lookup_oct8.html was obtained from the public stock lookup with month October, day 8 and year 2026. A later row in the response is outside the cutoff and is excluded.'),para('Provider crosscheck','BlackHead'),para('Yahoo MSFT original response<br/>Raw close 522.6099853515625<br/>Quoted valuation close rounded to cents <b>522.61</b><br/>Daily return sample joins MSFT and SPY adjusted closes by date.'),para('Source URL https://microsoft.gcs-web.com/'),PageBreak(),para('Treasury and equity premium','BlackTitle'),para('US Treasury par yield curve','BlackHead'),para('Selected row 10/08/2026<br/>Selected column 10 Yr<br/><b>Yield 5.22 percent</b><br/>Decimal model input 0.0522'),para('Source https://home.treasury.gov/resource-center/data-chart-center/interest-rates/<br/>Original files treasury_oct2026.csv and treasury_oct2026.html'),para('Damodaran equity risk premium','BlackHead'),para('Explicit publication date October 1 2026<br/>Trailing twelve months with adjusted payout<br/><b>Implied ERP 4.20 percent</b><br/>Decimal model input 0.042'),para('Original file damodaran_home.html and October calculator ERPOct26.xlsx. The monthly archive last date serial conflicts with its stated period, so the author page determines the dated selection.'),para('Source https://pages.stern.nyu.edu/~adamodar/New_Home_Page/home.htm'),PageBreak(),para('Debt spread and inference','BlackTitle'),para('Synthetic credit spread','BlackHead'),para('Damodaran ratings table January 2026<br/>Large non-financial company AAA row<br/><b>Default spread 0.40 percent</b><br/>Added to the cutoff risk-free rate as a historical proxy.'),para('Original file damodaran_ratings_jan2026.html. This is not a same-day Microsoft bond yield or directly observed lease discount rate.'),para('HC3 calculation method','BlackHead'),para('The regression uses daily MSFT and SPY total-return proxies with an intercept. HC3 scales each squared residual by 1 / (1 - leverage) squared in the covariance sandwich. Beta stays equal to the OLS point estimate. Newey-West lag five is shown separately to check serial dependence.'),para('Method source https://www.statsmodels.org/stable/generated/<br/>statsmodels.regression.linear_model.OLSResults.HC3_se.html<br/>Original file hc3_method.html. The delivered Python implements the matrix calculation with NumPy and SciPy.')]
    doc_pdf('Market_Source_Extracts.pdf',story)
    p=ROOT/'deliverables/Market_Source_Extracts.pdf'
    d=fitz.open(p)
    for page in d:
        for token in ['USD 522.61','522.61','5.22 percent','4.20 percent','0.40 percent']:
            for rect in page.search_for(token):
                a=page.add_highlight_annot(rect);a.set_colors(stroke=(1,.82,.20));a.set_info(title='Used market input',content='Generated source extract. Check the preserved original raw file for provenance.');a.update()
    d.save(p.with_name('_market_extracts.pdf'),garbage=2,deflate=True);d.close();p.with_name('_market_extracts.pdf').replace(p)

if __name__=='__main__':
    main()
