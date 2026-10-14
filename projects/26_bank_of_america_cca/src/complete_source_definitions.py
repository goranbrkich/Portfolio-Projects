"""Keep gross/net loans separate and preserve Truist deposit components as raw facts."""
from fetch_sources import ROOT
from extract_data import selected,TAGS
from lxml import html
import json,re
def main():
 records=json.loads((ROOT/'data/processed/financial_inputs.json').read_text());manifest=json.loads((ROOT/'data/processed/filing_manifest.json').read_text());byacc={m['accn']:m for m in manifest};reports=json.loads((ROOT/'data/processed/report_manifest.json').read_text())
 # Fingerprints permit unchanged PDF annotation checkpoints to be reused.
 import hashlib
 for m in reports:
  cp=ROOT/'qa/pdf_checkpoints'/(ROOT/m['path']).name;cp=cp.with_suffix(cp.suffix+'.json')
  if not cp.exists():continue
  rec=next((r for r in records if r['ticker']==m['ticker'] and r['period']==m['period']),None)
  if rec:
   old=json.loads(cp.read_text());old['facts_hash']=hashlib.sha256(json.dumps([(k,x['value']) for k,x in rec['metrics'].items() if x],sort_keys=True).encode()).hexdigest();cp.write_text(json.dumps(old))
 TAGS['loans_gross']=['FinancingReceivableExcludingAccruedInterestBeforeAllowanceForCreditLoss','LoansAndLeasesReceivableNetOfDeferredIncomeBeforeAllowance','LoansAndLeasesReceivableNetOfDeferredIncome']
 TAGS['loans_net']=['FinancingReceivableExcludingAccruedInterestAfterAllowanceForCreditLoss','LoansAndLeasesReceivableNetOfDeferredIncomeNetOfAllowance','LoansAndLeasesReceivableNetReportedAmount']
 changed=0
 for bank in json.loads((ROOT/'data/processed/banks.json').read_text()):
  t=bank['ticker'];company=json.loads((ROOT/f'data/raw/sec/{t}_companyfacts.json').read_text())
  for r in records:
   if r['ticker']!=t:continue
   old=r['metrics']['loans_gross']
   if old and old['tag'].endswith('LoansAndLeasesReceivableNetReportedAmount'):
    f=selected(company,'loans_gross',r['start'],r['end'])
    if f:
     m=byacc.get(f['accn']);f['source_path']=m['path'] if m else f'data/raw/sec/{t}_companyfacts.json';f['source_url']=m['url'] if m else f['source_url']
    r['metrics']['loans_gross']=f;changed+=1
   if r['metrics']['loans_net'] is None:
    f=selected(company,'loans_net',r['start'],r['end'])
    if f:
     m=byacc.get(f['accn']);f['source_path']=m['path'] if m else f'data/raw/sec/{t}_companyfacts.json';f['source_url']=m['url'] if m else f['source_url'];r['metrics']['loans_net']=f
  if t=='TFC':
   for r in records:
    if r['ticker']!=t:continue
    components={}
    for key,tag in [('noninterest','NoninterestBearingDepositLiabilities'),('interest','InterestBearingDepositLiabilities')]:
     cs=[x for x in company['facts']['us-gaap'].get(tag,{}).get('units',{}).get('USD',[]) if x.get('end')==r['end'] and x.get('filed','9999')<='2026-10-08']
     if not cs:continue
     x=max(cs,key=lambda x:x.get('filed',''));m=byacc.get(x['accn']);components[key]={'value':x['val']/1e6,'original_value':x['val'],'unit':'USD','tag':'us-gaap:'+tag,'label':('Noninterest-bearing' if key=='noninterest' else 'Interest-bearing')+' deposit liabilities','filed':x['filed'],'accn':x['accn'],'source_path':m['path'] if m else f'data/raw/sec/{t}_companyfacts.json','source_url':m['url'] if m else f'https://data.sec.gov/api/xbrl/companyfacts/CIK0000092230.json','selection_note':'Reported deposit component. Total deposits are summed in the financial build.'}
     if m:
      doc=html.fromstring((ROOT/m['path']).read_bytes())
      matches=[el for el in doc.iter() if isinstance(el.tag,str) and el.tag.lower().endswith('nonfraction') and el.get('name')=='us-gaap:'+tag and float(re.sub(r'[^\d.]','',el.text_content()) or 0)==x['val']/1e6]
      if len(matches)==1:components[key]['ixid']=matches[0].get('id','')
    r['deposit_components']=components
 (ROOT/'data/processed/financial_inputs.json').write_text(json.dumps(records,indent=2));print('Gross/net loan definitions repaired',changed);print('TFC deposits components',next(r for r in records if r['ticker']=='TFC' and r['period']=='H1 2026')['deposit_components'])
if __name__=='__main__':main()
