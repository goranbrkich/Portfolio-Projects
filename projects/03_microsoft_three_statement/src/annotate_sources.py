"""Link the workbook's source cells to native filing PDFs and annotate copies.

--plan-only reads the PDFs and writes the crosswalk before any PDF authoring.
Original reports and the Excel/Word deliverables are never modified.
"""
from __future__ import annotations
import argparse
from collections import defaultdict, Counter
import csv
import hashlib
import json
import math
from pathlib import Path
import re

import fitz

from workbook_io import PROJECT, read_workbook
from normalize_filings import norm, IS_ALIASES, BS_ALIASES, CF_ALIASES

COLORS = {'income': (0.35, 0.86, 0.50), 'balance': (0.32, 0.68, 0.96),
          'cashflow': (1.00, 0.82, 0.20), 'notes': (0.70, 0.48, 0.92), 'segments': (0.70, 0.48, 0.92)}
TITLES = {'income': 'Income statement', 'balance': 'Balance sheet', 'cashflow': 'Cash flow',
          'notes': 'Note disclosures', 'segments': 'Segments'}


def primary_labels(kind, rows, pos):
    result = {}
    section = ''
    for label, values in rows:
        key = norm(label)
        if not values:
            if key.startswith('current assets'): section = 'ca'
            elif key.startswith('current liabilities'): section = 'cl'
            elif key.startswith('earnings per share'): section = 'eps'
            elif key.startswith('weighted average shares'): section = 'shares'
            elif key == 'financing': section = 'financing'
            elif key == 'investing': section = 'investing'
            continue
        if pos >= len(values): continue
        value = values[pos]
        names = []
        if kind == 'income':
            names = [name for name, aliases in IS_ALIASES.items() if any(key.startswith(a) for a in aliases)]
            if key in ('basic', 'diluted') and section in ('eps', 'shares'):
                names.append(f'{section}_{key}')
        elif kind == 'balance':
            names = [name for name, alias in BS_ALIASES.items() if (key == alias if name.startswith('total_') else key.startswith(alias))]
            if key == 'equity investments': names.append('equity_investments')
            if key in ('other', 'other current assets') and section == 'ca': names.append('other_current_assets')
            if key in ('other', 'other current liabilities') and section == 'cl': names.append('other_current_liab')
            if key == 'income taxes' and section == 'cl': names.append('current_tax')
        else:
            names = [name for name, alias in CF_ALIASES.items() if key.startswith(alias) and not (name == 'debt_issued' and ('maturities' in key or 'short term' in key))]
            if 'goodwill and asset impairments' in key: names.append('impairments')
            if key.startswith('net cash') and 'financing' in key: names.append('cff')
            if key.startswith('net cash') and 'investing' in key: names.append('cfi')
            if ('maturities of 90 days or less' in key or 'short term debt' in key) and section == 'financing': names.append('net_st_debt')
            if key.startswith('cash premium on debt exchange'): names.append('debt_exchange_cash_premium')
            if key in ('other', 'other net') and section in ('financing', 'investing'):
                names.append('other_financing' if section == 'financing' else 'other_investing')
            if key == 'securities lending payable': names.append('securities_lending_cf')
        for name in names:
            result[name] = {'label': label, 'source_value': value, 'section': section}
    return result


def number(token):
    t = token.replace(',', '').replace('$', '').replace('\u2212', '-').strip()
    if t in ('-', '\u2013', '\u2014'): return 0.0
    if not re.fullmatch(r'\(?[+-]?\d+(?:\.\d+)?\)?%?', t): return None
    negative = t.startswith('(')
    percent = t.endswith('%')
    v = float(t.rstrip('%').strip('()'))
    return (-abs(v) if negative else v) / (100 if percent else 1)


def read_pages(path):
    pages = []
    with fitz.open(path) as doc:
        for page in doc:
            words = page.get_text('words')
            lines = []
            for w in sorted(words, key=lambda w: ((w[1]+w[3])/2, w[0])):
                y = (w[1]+w[3])/2
                line = next((ln for ln in lines[-4:] if abs(ln['y']-y)<2.4), None)
                if line is None:
                    line = {'y': y, 'words': []}; lines.append(line)
                line['words'].append(w)
            for line in lines:
                line['words'].sort(key=lambda w:w[0])
                line['text'] = ' '.join(w[4] for w in line['words'])
            pages.append({'text': page.get_text(), 'words': words, 'lines': lines,
                          'width': page.rect.width, 'height': page.rect.height})
    return pages


def scopes(pages):
    found = {}
    tests = {'income': 'INCOME STATEMENTS', 'balance': 'BALANCE SHEETS', 'cashflow': 'CASH FLOWS STATEMENTS'}
    for kind, heading in tests.items():
        matches = [i for i,p in enumerate(pages) if heading in p['text'][:700]
                   and not (kind=='income' and 'COMPREHENSIVE INCOME STATEMENTS' in p['text'][:700])]
        if not matches: raise ValueError(f'Missing {heading}')
        found[kind] = [matches[0]]
        # A cash-flow statement can continue on the next page.
        if kind=='cashflow' and 'Cash and cash equivalents, end of period' not in pages[matches[0]]['text']:
            found[kind].append(matches[0]+1)
    note_start = next((i for i,p in enumerate(pages) if re.search(r'NOTE\s+1\s*[-\u2013\u2014]',p['text'][:1500])), found['cashflow'][-1]+2)
    note_end = next((i for i in range(note_start+1,len(pages)) if 'REPORT OF INDEPENDENT' in pages[i]['text'][:800] or 'ITEM 9.' in pages[i]['text'][:500]), min(len(pages),note_start+50))
    found['notes'] = list(range(note_start,note_end))
    found['segments'] = [i for i in found['notes'] if 'Productivity and Business Processes' in pages[i]['text'] and 'Intelligent Cloud' in pages[i]['text'] and ('Operating income' in pages[i]['text'] or 'operating income' in pages[i]['text'])]
    return found


def tokens(label):
    return set(norm(label).split()) - {'the','and','of','in','for','to','total','net','other','expense','income','lease','value'}


def locate(record, pages, scope):
    target = record['source_value']
    if target is None:
        return {'page': scope[record['category']][0]+1, 'rects': [], 'status': record['presentation'], 'source_text': ''}
    indices = scope[record['category']]
    if record['metric']=='construction_commitments':
        indices=[i for i,p in enumerate(pages) if 'Construction commitments' in p['text'] and 'In millions' in p['text']]
    wanted = tokens(record['source_label'])
    candidates = []
    for pi in indices:
        p=pages[pi]
        year_words=[w for w in p['words'] if w[4]==str(record['fiscal_year']) and w[0]>p['width']*.45 and w[1]<200]
        year_x=min(year_words,key=lambda w:w[1])[2] if year_words and record['category'] in ('income','balance','cashflow') else None
        for line in p['lines']:
            label_words=[w for w in line['words'] if w[0]<p['width']*.64]
            label_text=' '.join(w[4] for w in label_words)
            score=len(wanted & set(norm(label_text).split()))/(len(wanted) or 1)
            # Include nearby prose for disclosed quantities expressed in billions.
            nearby=' '.join(ln['text'] for ln in p['lines'] if abs(ln['y']-line['y'])<24)
            prose_score=len(wanted & set(norm(nearby).split()))/(len(wanted) or 1)
            for w in line['words']:
                v=number(w[4])
                if v is None:continue
                scale=1; exact=math.isclose(v,target,rel_tol=1e-9,abs_tol=1e-6)
                if not exact and record['category']=='notes' and 'billion' in nearby.lower() and math.isclose(v*1000,target,rel_tol=1e-9,abs_tol=.01):
                    scale=1000;exact=True
                if not exact:continue
                if year_x is not None and abs(w[2]-year_x)>14:continue
                s=score*8+prose_score*2
                if str(record['fiscal_year']) in nearby:s+=1
                if record['category'] in ('income','balance','cashflow'):
                    if year_x is None:continue
                    s+=4
                if record['category']=='segments':
                    seg=record.get('segment','')
                    # Segment headings can precede a metric row in the detailed tables.
                    prev=[ln['text'] for ln in p['lines'] if 0<=line['y']-ln['y']<125]
                    segs=[ln for ln in prev if any(name in ln for name in ['Productivity and Business Processes','Intelligent Cloud','More Personal Computing'])]
                    if segs and seg in segs[-1]:s+=6
                    if seg in label_text:s+=8
                    if 'revenue' in record['metric'] and 'Revenue' in nearby:s+=2
                rects=[list(w[:4])]
                # Also highlight the matching row label, excluding other numeric columns.
                if score>=.45:
                    for lw in label_words:
                        if number(lw[4]) is None and lw[0]<w[0]-15:rects.append(list(lw[:4]))
                candidates.append((s,pi,w,rects,nearby,scale))
    if candidates:
        best=max(candidates,key=lambda x:x[0])
        status='reported_billion_to_million' if best[5]==1000 else record['presentation']
        return {'page':best[1]+1,'rects':best[3],'status':status,'source_text':best[4],'scale':best[5],'match_score':round(best[0],3)}
    # Explicitly retain the evidence gap, never substitute an unrelated number.
    return {'page':indices[0]+1 if indices else None,'rects':[], 'status':'needs_review','source_text':''}


def catalog(data,layout,wb,primary,facts):
    records=[];columns={y:c for y,c in zip(range(2017,2027),'EFGHIJKLMN')}
    tables={(t['report_year'],t['statement'],t['table_index']):t for t in primary}
    statement_sheets={'income':('Income Statement','IS'),'balance':('Balance Sheet','BS'),'cashflow':('Cash Flow','CF')}
    for y in range(2017,2027):
        for kind,mapping in layout['raw'].items():
            prov=data['provenance'][f'{y}.{kind}'];table=tables[(prov['report_year'],kind,prov['table_index'])]
            labels=primary_labels(kind,table['rows'],table['years'].index(y));cell_col=columns[y]
            for metric,row in mapping.items():
                value=data['actuals'][str(y)][kind][metric];source=labels.get(metric)
                cells=[f'Historical Inputs!{cell_col}{row}']
                sn,sk=statement_sheets[kind]
                if metric in layout[sk]:cells.append(f'{sn}!{cell_col}{layout[sk][metric]}')
                if metric in ('revenue','cfo'):
                    if kind=='income' and metric=='revenue':cells.append(f'Financial Modeling!D{y-2017+12}')
                    if kind=='cashflow' and metric=='cfo':cells.append(f'Financial Modeling!E{y-2017+12}')
                presentation=('calculated_residual' if kind=='cashflow' and metric=='other_operating_adjustments'
                              else 'reported' if source else 'normalized_absence' if value==0 else 'calculated_residual')
                parts=[source] if source else [{'label':metric.replace('_',' '),'source_value':None}]
                if kind=='cashflow' and metric=='other_financing' and data['actuals'][str(y)][kind]['debt_exchange_cash_premium']:
                    parts.append(labels['debt_exchange_cash_premium']);presentation='calculated_aggregate'
                for part in parts:
                    records.append({'category':kind,'metric':metric,'fiscal_year':y,'report_year':prov['report_year'],
                                    'source_url':prov['url'],'source_label':part['label'],'source_value':part['source_value'],
                                    'workbook_value':value,'unit':'USD per share' if metric.startswith('eps') else 'million shares' if metric.startswith('shares') else 'USD millions',
                                    'cells':cells,'presentation':presentation,'source_table_index':prov['table_index']})
        for metric,row in layout['nr'].items():
            if metric not in data['notes'][str(y)]:continue
            fact=next((f for f in facts if f['year']==y and f['metric']==metric),None)
            if fact is None:raise ValueError(f'Missing frozen fact: {y} {metric}')
            magnitude = metric in ('accumulated_depreciation','interest_expense')
            source_label={'ppe_gross':'Total at cost','ppe_depreciation':'Depreciation expense',
                          'intangible_amortization':'Amortization expense',
                          'finance_lease_principal':'Financing cash flows from finance leases',
                          'operating_lease_cost':'Operating lease cost'}.get(metric,metric.replace('_',' '))
            records.append({'category':'notes','metric':metric,'fiscal_year':y,'report_year':int(fact['fy']),
                            'source_url':f'https://www.sec.gov/Archives/edgar/data/789019/{fact["accn"].replace("-","")}/',
                            'source_label':source_label,'source_value':-data['notes'][str(y)][metric] if magnitude else data['notes'][str(y)][metric],
                            'workbook_value':data['notes'][str(y)][metric],'unit':'USD millions',
                            'cells':[f'Filing Notes!{columns[y]}{row}'],'presentation':'reported_magnitude' if magnitude else 'reported','xbrl_tag':fact['tag'],'accession':fact['accn']})
        seg=data['segments'][str(y)]
        for name,values in seg['data'].items():
            start={'Productivity and Business Processes':8,'Intelligent Cloud':22,'More Personal Computing':36}[name]
            for metric,value in values.items():
                # The Excel row is allocated forecast expense, with n.a. in actuals.
                # The reported historical expense is not a direct input in that row.
                if metric == 'operating_expenses': continue
                row={'revenue':start,'cost_of_revenue':start+6,'operating_expenses':start+9,'operating_income':start+10}[metric]
                records.append({'category':'segments','segment':name,'metric':metric,'fiscal_year':y,'report_year':seg['report_year'],
                                'source_url':seg['url'],'source_label':name+' '+metric.replace('_',' '),'source_value':value,
                                'workbook_value':value,'unit':'USD millions','cells':[f'Segments!{columns[y]}{row}'],'presentation':'reported'})
    for metric,row in layout['fixed'].items():
        v=data['notes']['2026'][metric]
        magnitude=metric in ('debt_discount','debt_hedge_adjustment','debt_exchange_premium')
        source_label={'debt_discount':'Unamortized discount and issuance costs',
                      'debt_hedge_adjustment':'Hedge fair value adjustments',
                      'debt_exchange_premium':'Premium on debt exchange',
                      'debt_face':'Total face value'}.get(metric,metric.replace('_',' '))
        records.append({'category':'notes','metric':metric,'fiscal_year':2026,'report_year':2026,'source_url':data['provenance']['2026.income']['url'],
                        'source_label':source_label,'source_value':-v if magnitude else v,'workbook_value':v,'unit':'USD millions',
                        'cells':[f'Filing Notes!N{row}'],'presentation':'reported_magnitude' if magnitude else 'reported'})
    for key,row in [('debt_maturities',37),('intangible_amortization_schedule',38),('op_lease_payments',39),('finance_lease_payments',40)]:
        for y,c,v in zip(range(2027,2032),'OPQRS',data['notes']['2026'][key]):
            records.append({'category':'notes','metric':key,'fiscal_year':y,'report_year':2026,'source_url':data['provenance']['2026.income']['url'],
                            'source_label':str(y)+' '+key.replace('_',' '),'source_value':v,'workbook_value':v,'unit':'USD millions',
                            'cells':[f'Filing Notes!{c}{row}'],'presentation':'reported_contractual_schedule'})
    records.append({'category':'notes','metric':'debt_thereafter','fiscal_year':2026,'report_year':2026,'source_url':data['provenance']['2026.income']['url'],
                    'source_label':'Thereafter debt maturities','source_value':data['notes']['2026']['debt_thereafter'],
                    'workbook_value':data['notes']['2026']['debt_thereafter'],'unit':'USD millions','cells':['Filing Notes!N42'],'presentation':'reported_contractual_schedule'})
    # FY2016 is opening-balance context, not an eleventh historical fiscal year.
    opening=data['actuals']['2016']['balance']['cash'];prov=data['provenance']['2016.balance']
    records.append({'category':'balance','metric':'opening_cash_context','fiscal_year':2016,'report_year':prov['report_year'],
                    'source_url':prov['url'],'source_label':'Cash and cash equivalents','source_value':opening,
                    'workbook_value':opening,'unit':'USD millions','cells':['Historical Inputs!E98','Cash Flow!E53'],
                    'presentation':'opening_balance_context'})
    # Check the actual source inputs before creating annotations.
    for i,r in enumerate(records,1):
        r['id']=f'SRC-{i:04}'
        sn,cell=r['cells'][0].split('!');actual=wb[sn].get(cell,{}).get('value')
        if not isinstance(actual,(int,float)) or not math.isclose(actual,r['workbook_value'],abs_tol=1e-6):
            raise ValueError(f'Source cell mismatch: {sn}!{cell}: {actual} vs {r["workbook_value"]}')
    return records


def make_plan():
    data=json.loads((PROJECT/'data/processed/historical_financials.json').read_text());layout=json.loads((PROJECT/'evidence/workbook_layout.json').read_text())
    primary=json.loads((PROJECT/'data/raw/primary_statement_tables.json').read_text());facts=json.loads((PROJECT/'data/raw/selected_sec_facts.json').read_text())
    wb,_=read_workbook(PROJECT/'deliverables/Microsoft_Three_Statement_Model.xlsx')
    records=catalog(data,layout,wb,primary,facts)
    pages={};scope={}
    for y in range(2017,2027):
        pages[y]=read_pages(PROJECT/f'data/raw/reports/originals/Microsoft_FY{y}_10K.pdf');scope[y]=scopes(pages[y])
    for r in records:
        r.update(locate(r,pages[r['report_year']],scope[r['report_year']]))
        r['original_pdf']=f'data/raw/reports/originals/Microsoft_FY{r["report_year"]}_10K.pdf'
        r['highlighted_pdf']=f'data/raw/reports/highlighted/Microsoft_FY{r["report_year"]}_10K_highlighted.pdf'
    plan={'color_legend':{k:list(v) for k,v in COLORS.items()},'source_records':records,'source_scopes':scope,
          'report_pages':{str(y):len(p) for y,p in pages.items()},'model_cutoff':'2026-10-08'}
    (PROJECT/'evidence/raw_materials_mapping.json').write_text(json.dumps(plan,indent=2)+'\n')
    print(json.dumps({'records':len(records),'status_counts':dict(Counter(r['status'] for r in records)),
                      'needs_review':[(r['id'],r['report_year'],r['category'],r['metric'],r['fiscal_year'],r['source_value']) for r in records if r['status']=='needs_review']},indent=2))
    return plan


def write_crosswalk(plan):
    fields=['source_id','category','metric','fiscal_year','report_year','source_label','source_value','reported_pdf_value','reported_pdf_unit','workbook_value','unit','conversion','excel_cells','original_pdf','highlighted_pdf','pdf_page','status','source_url','xbrl_tag','accession']
    with (PROJECT/'docs/source_crosswalk.csv').open('w',newline='',encoding='utf-8-sig') as f:
        writer=csv.DictWriter(f,fieldnames=fields);writer.writeheader()
        for r in plan['source_records']:
            writer.writerow({'source_id':r['id'],'category':r['category'],'metric':r['metric'],'fiscal_year':r['fiscal_year'],
                             'report_year':r['report_year'],'source_label':r['source_label'],'source_value':r['source_value'],
                             'reported_pdf_value':r['source_value']/r.get('scale',1) if r['source_value'] is not None else None,
                             'reported_pdf_unit':'USD billions' if r.get('scale')==1000 else r['unit'],
                             'workbook_value':r['workbook_value'],'unit':r['unit'],
                             'conversion':'multiply by 1000' if r.get('scale')==1000 else 'absolute magnitude' if r['presentation']=='reported_magnitude' else r['presentation'] if r['source_value'] is None or r['presentation']=='calculated_aggregate' else 'none',
                             'excel_cells':' | '.join(r['cells']),
                             'original_pdf':r['original_pdf'],'highlighted_pdf':r['highlighted_pdf'],'pdf_page':r['page'],'status':r['status'],
                             'source_url':r['source_url'],'xbrl_tag':r.get('xbrl_tag',''),'accession':r.get('accession','')})


def annotate(plan,resume=False):
    records=plan['source_records'];results=[]
    if any(r['status']=='needs_review' for r in records):raise ValueError('Review unmatched records before authoring PDFs')
    for year in range(2017,2027):
        original=PROJECT/f'data/raw/reports/originals/Microsoft_FY{year}_10K.pdf'
        output=PROJECT/f'data/raw/reports/highlighted/Microsoft_FY{year}_10K_highlighted.pdf';output.parent.mkdir(parents=True,exist_ok=True)
        grouped=defaultdict(list)
        for r in records:
            if r['report_year']==year and r['page']:grouped[r['page']].append(r)
        if resume and output.exists() and output.stat().st_size:
            with fitz.open(original) as a,fitz.open(output) as b:
                if len(a)!=len(b) or any(a[i].get_text()!=b[i].get_text() for i in range(len(a))):raise ValueError(f'Invalid resume file: {year}')
                annotations=sum(1 for page in b for annot in page.annots() or [] if annot.type[0]==8)
            results.append({'report_year':year,'used_pdf_pages':sorted(grouped),'value_and_label_annotations':annotations,
                            'source_text_and_page_count_preserved':True,'original_sha256':hashlib.sha256(original.read_bytes()).hexdigest(),
                            'highlighted_sha256':hashlib.sha256(output.read_bytes()).hexdigest()})
            continue
        doc=fitz.open(original)
        annotations=0
        for page_num,rs in sorted(grouped.items()):
            page=doc[page_num-1];marked=defaultdict(list)
            for r in rs:
                for rect in r['rects']:
                    marked[(r['category'],tuple(round(v,2) for v in rect))].append(r)
            for (category,rect),linked in marked.items():
                annot=page.add_highlight_annot(fitz.Rect(rect));annot.set_colors(stroke=COLORS[category]);annot.set_opacity(.32)
                body='\n'.join(f'{r["id"]}: FY{r["fiscal_year"]} {r["metric"]}; source {r["source_value"]}; Excel {r["workbook_value"]} {r["unit"]}; '+', '.join(r['cells'])+f'; {r["status"]}' for r in linked)
                annot.set_info(title='Goran Brkich Portfolio',subject=TITLES[category],content=body);annot.update();annotations+=1
            categories=sorted({r['category'] for r in rs});dominant=Counter(r['category'] for r in rs).most_common(1)[0][0]
            # A margin stripe marks the entire used page without covering source text.
            stripe=page.add_rect_annot(fitz.Rect(10,24,14,page.rect.height-24));stripe.set_colors(stroke=COLORS[dominant],fill=COLORS[dominant]);stripe.set_opacity(.55)
            stripe.set_info(title='Page used in Excel',content='; '.join(TITLES[c] for c in categories)+'\n'+', '.join(r['id'] for r in rs));stripe.update()
        doc.set_toc([[1,f'Used in Excel - PDF page {p}: '+', '.join(sorted({TITLES[r["category"]] for r in rs})),p] for p,rs in sorted(grouped.items())])
        metadata=doc.metadata;metadata['subject']='Review copy: source figures used in Microsoft Three Statement Model. Green income; blue balance; yellow cash flow; purple notes/segments. Original pagination retained.';doc.set_metadata(metadata)
        # Avoid quadratic stream deduplication on issuer PDFs with many objects.
        doc.save(output,garbage=2,deflate=True,use_objstms=1);doc.close()
        with fitz.open(original) as a,fitz.open(output) as b:
            if len(a)!=len(b) or any(a[i].get_text()!=b[i].get_text() for i in range(len(a))):raise ValueError(f'Source text/pages changed: {year}')
        results.append({'report_year':year,'used_pdf_pages':sorted(grouped),'value_and_label_annotations':annotations,
                        'source_text_and_page_count_preserved':True,'original_sha256':hashlib.sha256(original.read_bytes()).hexdigest(),
                        'highlighted_sha256':hashlib.sha256(output.read_bytes()).hexdigest()})
    write_crosswalk(plan)
    result={'reports':results,'source_records':len(records),'status_counts':dict(Counter(r['status'] for r in records)),
            'original_workbook_sha256':hashlib.sha256((PROJECT/'deliverables/Microsoft_Three_Statement_Model.xlsx').read_bytes()).hexdigest(),
            'model_cells_modified':False,'original_report_bytes_preserved':True}
    (PROJECT/'evidence/raw_materials_verification.json').write_text(json.dumps(result,indent=2)+'\n');print(json.dumps(result,indent=2))


if __name__=='__main__':
    parser=argparse.ArgumentParser(description=__doc__);parser.add_argument('--plan-only',action='store_true');parser.add_argument('--resume',action='store_true');args=parser.parse_args()
    if args.plan_only:make_plan()
    else:annotate(json.loads((PROJECT/'evidence/raw_materials_mapping.json').read_text()),resume=args.resume)
