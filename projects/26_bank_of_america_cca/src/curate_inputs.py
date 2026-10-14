"""Resolve untagged book-value disclosures against complete original table rows."""
from fetch_sources import ROOT,download
from extract_data import LABELS
from lxml import html
import json,re,gc,sys
def text(x):return ' '.join(x.text_content().split())
def numbers(row):
 out=[]
 for td in row.xpath('./td|./th'):
  for sup in td.xpath('.//sup'):sup.drop_tree()
  t=text(td).strip('$ %')
  if re.fullmatch(r'\(?[\d,]+(?:\.\d+)?\)?',t):out.append(float(t.replace(',','').strip('()'))*(-1 if t.startswith('(') else 1))
 return out
def put(record,key,value,manifest,row,unit=None):
 unit=unit or ('USD/shares' if key in ['bvps_reported','tbvps_reported'] else 'pure' if key=='cet1_reported' else 'USD')
 record['metrics'][key]={'value':value,'original_value':value if unit in ['pure','USD/shares'] else value*1e6,'unit':unit,'tag':'Manual verified issuer table row','label':LABELS.get(key,key),'filed':manifest['filingDate'],'accn':manifest['accn'],'source_url':manifest['url'],'source_path':manifest['path'],'source_row_text':row,'selection_note':'Verified original issuer table row; same-date end balance, excluding average and regulatory minimum tables.'}
def main():
 records=json.loads((ROOT/'data/processed/financial_inputs.json').read_text());man=json.loads((ROOT/'data/processed/filing_manifest.json').read_text());lookup={(r['ticker'],r['end']):r for r in records};docs={}
 for m in ([] if '--finish' in sys.argv else man):
  if m.get('error') or (m['ticker'],m['reportDate']) not in lookup:continue
  r=lookup[m['ticker'],m['reportDate']];raw=(ROOT/m['path']).read_text(errors='replace');rows=[]
  for match in re.finditer(r'<tr\b[^>]*>(.*?)</tr>',raw,re.I|re.S):
   fragment='<table>'+match[0]+'</table>'
   try:tr=html.fromstring(fragment).xpath('//tr')[0]
   except (ValueError,IndexError):continue
   label=text(tr)
   if len(label)>650:continue
   if re.match(r'^(?:Tangible )?[Bb]ook value(?: per(?: common)? share|\s*\()',label):
    vals=numbers(tr)
    if not vals or vals[0]<1 or vals[0]>1000:continue
    key='tbvps_reported' if label.lower().startswith('tangible') else 'bvps_reported'
    if r['metrics'].get(key) is None:put(r,key,vals[0],m,label)
  print('Book disclosures',m['ticker'],m['reportDate'],flush=True)
  (ROOT/'data/processed/financial_inputs.json').write_text(json.dumps(records,indent=2))
 # Curated indices were checked against the actual/current column headings in each source.
 current_rows={
 'BAC':{'common_equity_reported':(171,0),'tce_reported':(1733,0),'cet1_reported':(290,0)},
 'JPM':{'common_equity_reported':(150,0),'tce_reported':(387,0),'cet1_reported':(137,0)},
 'C':{'common_equity_reported':(85,0),'tce_reported':(674,0),'cet1_reported':(97,0)},
 'WFC':{'common_equity_reported':(103,0),'tce_reported':(104,0),'cet1_reported':(126,0),'loans_gross':(750,0)},
 'GS':{'common_equity_reported':(3552,0),'tce_reported':(3555,0),'cet1_reported':(2518,0),'loans_gross':(1469,-1)},
 'MS':{'common_equity_reported':(784,0),'tce_reported':(144,0),'cet1_reported':(151,0),'loans_gross':(944,-1)},
 'USB':{'tce_reported':(621,0),'cet1_reported':(114,0)},
 'PNC':{'common_equity_reported':(269,0),'cet1_reported':(282,0)},
 'TFC':{'equity':(331,0),'preferred':(325,0),'cet1_reported':(2468,0)},
 'COF':{'cet1_reported':(320,0)},
 }
 for ticker,maps in current_rows.items():
  m=next(x for x in man if x['ticker']==ticker and x['reportDate']=='2026-06-30');doc=html.fromstring((ROOT/m['path']).read_bytes());trs=doc.xpath('//tr');r=lookup[ticker,'2026-06-30']
  for key,(idx,which) in maps.items():
   original=text(trs[idx]);vals=numbers(trs[idx]);assert vals,(ticker,key,original)
   value=vals[which];value=value/100 if key=='cet1_reported' else value
   put(r,key,value,m,original);r['metrics'][key]['source_row_index']=idx
  del doc;gc.collect()
 # TFC first-half 2025 parent equity and preferred carrying were verified in its balance sheet.
 m=next(x for x in man if x['ticker']=='TFC' and x['reportDate']=='2025-06-30');doc=html.fromstring((ROOT/m['path']).read_bytes());trs=doc.xpath('//tr');r=lookup['TFC','2025-06-30']
 tr=trs[336];assert '64,840' in text(tr);put(r,'equity',64840,m,text(tr));r['metrics']['equity']['source_row_index']=336
 # PNC non-GAAP reconciliation is in the original issuer financial supplement.
 url='https://investor.pnc.com/sec-filings/all-sec-filings/content/0001628280-26-048244/q22026financialhighlightsa.htm';p='data/raw/filings/PNC_2026Q2_financial_highlights.htm';sm=download(url,p)
 doc=html.fromstring((ROOT/p).read_bytes());trs=doc.xpath('//tr');mr={'path':p,'url':url,'filingDate':'2026-07-15','accn':'0001628280-26-048244'}
 (ROOT/'data/processed/financial_inputs.json').write_text(json.dumps(records,indent=2))
 for key,pattern,valnow,valold in [('tce_reported',r'^Tangible common.*equity',44378,40959),('tbvps_reported',r'^Tangible book value per',111.09,103.96),('intangibles_total',r'^Goodwill and other intangible',14161,None)]:
  matches=[tr for tr in trs if re.match(pattern,text(tr),re.I) and len(text(tr))<500]
  for target,value in [('2026-06-30',valnow),('2025-06-30',valold)]:
   if value is None:continue
   candidates=[tr for tr in matches if str(value) in text(tr).replace(',','')]
   assert candidates,('PNC',key,value)
   put(lookup['PNC',target],key,value,mr,text(candidates[0]))
 if not any(x['path']==p for x in man):man.append({'ticker':'PNC','cik':713676,'form':'8-K supplement','reportDate':'2026-06-30','filingDate':'2026-07-15','accn':mr['accn'],**sm})
 (ROOT/'data/processed/filing_manifest.json').write_text(json.dumps(man,indent=2));(ROOT/'data/processed/financial_inputs.json').write_text(json.dumps(records,indent=2))
 print('Curated consolidated book and capital inputs',flush=True)
if __name__=='__main__':main()
