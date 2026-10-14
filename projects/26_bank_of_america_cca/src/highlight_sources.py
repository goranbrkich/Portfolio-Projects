"""Preserve raw originals and create separate full HTML/PDF marked review copies.
Poppler extracts PDF word coordinates; pypdf writes annotations without MuPDF.
"""
from fetch_sources import ROOT
from pathlib import Path
from lxml import html,etree
from pypdf import PdfReader,PdfWriter
from pypdf.annotations import Highlight
from pypdf.generic import ArrayObject,FloatObject,NameObject,TextStringObject
import json,re,csv,collections,hashlib,subprocess,gc,sys
records=json.loads((ROOT/'data/processed/financial_inputs.json').read_text())
facts=[];row=7
components=[]
for r in records:
 for k,x in r['metrics'].items():
  if x:facts.append({'ticker':r['ticker'],'period':r['period'],'metric':k,'excel_input_cell':'Historical Inputs!F'+str(row),**x})
  if k=='deposits':
   for key,c in [('noninterest','T'),('interest','U')]:
    if key in r.get('deposit_components',{}):components.append({'ticker':r['ticker'],'period':r['period'],'metric':'deposits_'+key,'excel_input_cell':'Historical Inputs!'+c+str(row),**r['deposit_components'][key]})
  row+=1
facts.extend(components)
colors={'balance':('#ACD3F2','ACD3F2'),'income':('#B8DFB5','B8DFB5'),'cash':('#FFE99A','FFE99A'),'notes':('#D9C8EA','D9C8EA')}
flows={'nii','noninterest_income','revenue_reported','opex','provision','pretax','tax','net_income','common_income','preferred_div','eps','shares_diluted','chargeoffs','recoveries'}
def category(f):return 'cash' if f['metric'] in ['dividends','buybacks','dps'] else 'income' if f['metric'] in flows else 'notes' if f['metric'] in ['cet1_reported','rwa','nim_reported'] else 'balance'
def norm(t):return ' '.join(t.split())
def parse_html(data):return html.fromstring(data,parser=html.HTMLParser(encoding='utf-8'))
groups=collections.defaultdict(list)
for f in facts:groups[f.get('source_path','')].append(f)
htmlchecks=[];checkpoint=ROOT/'qa/html_highlight_checkpoint.json'
if '--resume-pdf' in sys.argv and checkpoint.exists():
 c=json.loads(checkpoint.read_text());facts=c['facts'];htmlchecks=c['html']
else:
 manifests=json.loads((ROOT/'data/processed/filing_manifest.json').read_text())
 for m in manifests:
  if m.get('error'):continue
  original=ROOT/m['path'];doc=parse_html(original.read_bytes());before=norm(''.join(doc.itertext()));nodes={x.get('id'):x for x in doc.xpath('//*[@id]')};trs=doc.xpath('//tr');exact=0
  for f in groups[m['path']]:
   node=nodes.get(f.get('ixid'));status='Exact XBRL anchor'
   if node is None and f.get('source_row_index') is not None:
    i=f['source_row_index'];node=trs[i] if i<len(trs) else None;status='Verified original table row'
   if node is None and f.get('source_row_text'):
    matches=[tr for tr in trs if norm(tr.text_content())==f['source_row_text']]
    if len(matches)==1:node=matches[0];status='Exact selected original table text'
   if node is not None:
    color=colors[category(f)][0];node.set('style',node.get('style','')+';background-color:'+color+';outline:1px solid '+color+';');node.set('title',node.get('title','')+f"\n{f['excel_input_cell']} | {f['ticker']} {f['period']} {f['metric']} | {f['value']} {f['unit']} | {status}");f['html_highlight_path']='highlighted/'+original.name;f['html_status']=status;exact+=1
   else:f['html_status']='Structured source retained; exact visible anchor requires review'
  out=ROOT/'highlighted'/original.name;out.write_bytes(etree.tostring(doc,method='html',encoding='utf-8'))
  assert before==norm(''.join(parse_html(out.read_bytes()).itertext())),original.name
  htmlchecks.append({'original':m['path'],'marked':str(out.relative_to(ROOT)),'matched_facts':exact,'source_text_preserved':True});del doc;gc.collect()
  if len(htmlchecks)%20==0:print('HTML annotated',len(htmlchecks),flush=True)
 checkpoint.write_text(json.dumps({'html':htmlchecks,'facts':facts}))
patterns={
'assets':r'total assets','liabilities':r'total liabilities','equity':r'(total|shareholders|stockholders).*equity','preferred':r'preferred','common_equity_reported':r'common.*equity',
'goodwill':r'goodwill','intangibles':r'intangible','intangibles_total':r'goodwill|intangible','shares_end':r'(common|shares).*outstanding','shares_issued':r'issued','shares_treasury':r'treasury',
'deposits':r'deposits','deposits_interest':r'deposit','deposits_noninterest':r'deposit','loans_gross':r'loans|financing receivable','loans_net':r'loans|financing receivable','allowance':r'allowance|reserve','cash':r'cash','aoci':r'accumulated|comprehensive',
'nii':r'net interest','noninterest_income':r'non.?interest (income|revenue)','revenue_reported':r'(net|total).*revenue','opex':r'(non.?interest|total).*expense','provision':r'provision|credit loss','pretax':r'before.*tax|pre.?tax','tax':r'income tax','net_income':r'net (income|earnings)','common_income':r'common.*|net (income|earnings)','preferred_div':r'preferred','eps':r'diluted','shares_diluted':r'weighted|diluted','dps':r'dividend','dividends':r'dividend','buybacks':r'repurchase|buyback','chargeoffs':r'charge.?off|write.?off','recoveries':r'recover','cet1_reported':r'common equity tier|cet1','rwa':r'risk.weighted','bvps_reported':r'book value','tbvps_reported':r'tangible.*book','tce_reported':r'tangible.*equity'}
def targets(f):
 v=abs(f['value'])*(100 if f['unit']=='pure' else 1)
 if f['unit'] in ['pure','USD/shares']:return list(dict.fromkeys([f'{v:.1f}',f'{v:.2f}']))
 return [f'{v:,.0f}',f'{v:.0f}'] if v>=1 and abs(v-round(v))<.00001 else [f'{v:,.1f}',f'{v:.1f}']
def content_hash(page):
 contents=page.get_contents();return hashlib.sha256(contents.get_data() if contents else b'').hexdigest()
pdfchecks=[];reports=json.loads((ROOT/'data/processed/report_manifest.json').read_text());(ROOT/'qa/pdf_checkpoints').mkdir(exist_ok=True)
for m in reports:
 original=ROOT/m['path'];cp=ROOT/'qa/pdf_checkpoints'/(original.name+'.json')
 relevant=[f for f in facts if f['ticker']==m['ticker'] and f['period']==m['period']]
 facts_hash=hashlib.sha256(json.dumps([(f['metric'],f['value']) for f in relevant],sort_keys=True).encode()).hexdigest()
 if cp.exists() and json.loads(cp.read_text()).get('facts_hash')==facts_hash:
  saved=json.loads(cp.read_text());pdfchecks.append(saved['check']);bykey={(x['ticker'],x['period'],x['metric']):x['pdf_evidence'] for x in saved['facts']}
  for f in facts:
   if (f['ticker'],f['period'],f['metric']) in bykey:f.setdefault('pdf_evidence',[]).extend(bykey[f['ticker'],f['period'],f['metric']])
  continue
 rawbbox=subprocess.run(['pdftotext','-bbox',str(original),'-'],capture_output=True,check=True).stdout;tree=etree.fromstring(rawbbox,etree.XMLParser(recover=True));pages=tree.xpath('//*[local-name()="page"]');numeric=collections.defaultdict(list);rowwords=[]
 for i,p in enumerate(pages):
  lines=collections.defaultdict(list)
  for w in p.xpath('./*[local-name()="word"]'):
   word=(float(w.get('xMin')),float(w.get('yMin')),float(w.get('xMax')),float(w.get('yMax')),''.join(w.itertext()));lines[int(word[1]/3)].append(word);numeric[word[4].strip('$()% ')].append((i,word))
  rowwords.append(lines)
 reader=PdfReader(original);before=[content_hash(p) for p in reader.pages];assert len(pages)==len(reader.pages);writer=PdfWriter();writer.clone_document_from_reader(reader);used=set();count=0;updated=[]
 for f in relevant:
  matches=[]
  for token in targets(f):
   for i,w in numeric.get(token,[]):
    words=[x for bucket in range(int(w[1]/3)-3,int(w[1]/3)+4) for x in rowwords[i].get(bucket,[])];context=' '.join(x[4] for x in words if abs(x[1]-w[1])<9 and x[0]<=w[2]+20)
    if re.search(patterns.get(f['metric'],re.escape(f['metric'])),context,re.I):matches.append((i,w,context))
  if not 1<=len(matches)<=3:continue
  for i,w,context in matches:
   h=float(pages[i].get('height'));x0,y0,x1,y1=w[:4];bottom,top=h-y1,h-y0;quads=ArrayObject([FloatObject(z) for z in [x0,top,x1,top,x0,bottom,x1,bottom]])
   ann=Highlight(rect=(x0,bottom,x1,top),quad_points=quads,highlight_color=colors[category(f)][1],printing=True);ann[NameObject('/T')]=TextStringObject('Excel source '+f['ticker']);ann[NameObject('/Contents')]=TextStringObject(f"{f['excel_input_cell']} | {f['period']} {f['metric']}: {f['value']} {f['unit']}. Metric-row/value evidence. Review the date column against the selected XBRL/table source. {context}")
   writer.add_annotation(i,ann);used.add(i+1);count+=1
  evidence={'original':m['path'],'marked':'highlighted/'+original.name,'pages':sorted({x[0]+1 for x in matches}),'status':'Metric-row/value evidence; selected source and date column recorded separately'};f.setdefault('pdf_evidence',[]).append(evidence);updated.append({'ticker':f['ticker'],'period':f['period'],'metric':f['metric'],'pdf_evidence':[evidence]})
 for p in sorted(used):writer.add_outline_item('Used source page '+str(p),p-1)
 out=ROOT/'highlighted'/original.name
 with out.open('wb') as stream:writer.write(stream)
 after=PdfReader(out);assert len(after.pages)==len(reader.pages) and before==[content_hash(p) for p in after.pages],original.name
 check={'original':m['path'],'marked':str(out.relative_to(ROOT)),'pages':len(reader.pages),'used_pages':sorted(used),'native_highlights':count,'page_contents_preserved':True};pdfchecks.append(check);cp.write_text(json.dumps({'check':check,'facts':updated,'facts_hash':facts_hash}));print('PDF annotated',original.name,count,flush=True)
 del reader,writer,after,tree,numeric,rowwords;gc.collect()
bykey={(f['ticker'],f['period'],f['metric']):f for f in facts};cross=[];row=7
for r in records:
 for k,x in r['metrics'].items():
  f=bykey.get((r['ticker'],r['period'],k),{});e=f.get('pdf_evidence',[]);cross.append({'excel_input_cell':'Historical Inputs!F'+str(row),'bank':r['ticker'],'period':r['period'],'metric':k,'value':x['value'] if x else 'n.a.','unit':x['unit'] if x else '', 'filed':x.get('filed','') if x else '', 'accession':x.get('accn','') if x else '', 'raw_original':x.get('source_path','') if x else '', 'source_url':x.get('source_url','') if x else '', 'xbrl_anchor':x.get('ixid','') if x else '', 'html_highlight':f.get('html_highlight_path',''),'html_status':f.get('html_status','Structured fact retained' if x else 'Not reported / unavailable'),'pdf_originals':'; '.join(a['original'] for a in e),'pdf_pages':'; '.join(','.join(map(str,a['pages'])) for a in e),'marked_pdf':'; '.join(a['marked'] for a in e),'pdf_status':'Secondary metric-row/value evidence; review date column' if e else 'No unambiguous PDF match; exact source retained','scaling':'USD and share source units divided by 1000000; ratios and USD/share unscaled'});row+=1
for comp in components:
 f=bykey[comp['ticker'],comp['period'],comp['metric']];e=f.get('pdf_evidence',[]);cross.append({'excel_input_cell':f['excel_input_cell'],'bank':f['ticker'],'period':f['period'],'metric':f['metric'],'value':f['value'],'unit':f['unit'],'filed':f['filed'],'accession':f['accn'],'raw_original':f['source_path'],'source_url':f['source_url'],'xbrl_anchor':f.get('ixid',''),'html_highlight':f.get('html_highlight_path',''),'html_status':f.get('html_status','Structured deposit component retained'),'pdf_originals':'; '.join(a['original'] for a in e),'pdf_pages':'; '.join(','.join(map(str,a['pages'])) for a in e),'marked_pdf':'; '.join(a['marked'] for a in e),'pdf_status':'Secondary metric-row/value evidence' if e else 'Exact original component source retained','scaling':'USD source divided by 1000000. Total calculated only in the bank financial build.'})
(ROOT/'data/processed/highlight_crosswalk.json').write_text(json.dumps(cross,indent=2))
with (ROOT/'data/processed/source_crosswalk.csv').open('w',newline='') as stream:
 w=csv.DictWriter(stream,fieldnames=list(cross[0]));w.writeheader();w.writerows(cross)
(ROOT/'qa/source_highlight_checks.json').write_text(json.dumps({'html':htmlchecks,'pdf':pdfchecks,'source_cells':len(cross),'reported_facts':len(facts)},indent=2));print('SOURCE ANNOTATION COMPLETE',len(htmlchecks),len(pdfchecks),flush=True)
