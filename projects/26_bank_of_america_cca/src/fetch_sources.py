"""Retrieve frozen primary bank data and unadjusted market prices. Reuses local originals."""
from pathlib import Path
from urllib.request import Request,urlopen
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime,timezone
import json,csv,time,hashlib,argparse
ROOT=Path(__file__).resolve().parents[1]
CUTOFF='2026-10-08'
UA='GoranBrkichPortfolioResearch https://github.com/goranbrkich/Portfolio-Projects'
BANKS=[('BAC',70858,'Bank of America','Universal'),('JPM',19617,'JPMorgan Chase','Universal'),('C',831001,'Citigroup','Universal'),('WFC',72971,'Wells Fargo','Commercial'),('GS',886982,'Goldman Sachs','Markets and investment banking'),('MS',895421,'Morgan Stanley','Markets and wealth management'),('USB',36104,'U.S. Bancorp','Commercial'),('PNC',713676,'PNC Financial Services','Commercial'),('TFC',92230,'Truist Financial','Commercial'),('COF',927628,'Capital One Financial','Consumer and cards')]
def download(url,path):
 p=ROOT/path;p.parent.mkdir(parents=True,exist_ok=True)
 if not p.exists():
  for attempt in range(3):
   try:
    with urlopen(Request(url,headers={'User-Agent':UA,'Accept-Encoding':'identity'}),timeout=35) as response:data=response.read()
    if not data:raise ValueError('Empty response')
    p.write_bytes(data);break
   except Exception:
    if attempt==2:raise
    time.sleep(attempt+1)
 data=p.read_bytes()
 return {'path':path,'url':url,'bytes':len(data),'sha256':hashlib.sha256(data).hexdigest(),'retrieved_at':datetime.now(timezone.utc).isoformat()}
def initial(bank):
 ticker,cik,name,group=bank;out=[]
 for kind,url in [('companyfacts',f'https://data.sec.gov/api/xbrl/companyfacts/CIK{cik:010}.json'),('submissions',f'https://data.sec.gov/submissions/CIK{cik:010}.json')]:out.append(download(url,f'data/raw/sec/{ticker}_{kind}.json'))
 start=int(datetime(2016,1,1,tzinfo=timezone.utc).timestamp());end=int(datetime(2026,10,9,tzinfo=timezone.utc).timestamp())
 url=f'https://query1.finance.yahoo.com/v8/finance/chart/{ticker}?period1={start}&period2={end}&interval=1d'
 out.append(download(url,f'data/raw/market/{ticker}_daily_unadjusted.json'))
 data=json.loads((ROOT/out[-1]['path']).read_text())['chart']['result'][0];q=data['indicators']['quote'][0]['close'];prices=[]
 for ts,close in zip(data['timestamp'],q):
  date=datetime.fromtimestamp(ts,timezone.utc).strftime('%Y-%m-%d')
  if close is not None and date<=CUTOFF:prices.append({'date':date,'close':close})
 assert prices[-1]['date']==CUTOFF,(ticker,prices[-1])
 with (ROOT/f'data/processed/{ticker}_prices.csv').open('w',newline='') as stream:
  writer=csv.DictWriter(stream,fieldnames=['date','close']);writer.writeheader();writer.writerows(prices)
 return {'ticker':ticker,'cik':cik,'name':name,'group':group,'asof_close':prices[-1]['close'],'price_date':prices[-1]['date'],'year_end_prices':{str(y):[x['close'] for x in prices if x['date']<=f'{y}-12-31'][-1] for y in range(2016,2026)}},out
def filings():
 items=[]
 for ticker,cik,_,_ in BANKS:
  sub=json.loads((ROOT/f'data/raw/sec/{ticker}_submissions.json').read_text());blocks=[sub['filings']['recent']]
  for hist in sub['filings'].get('files',[]):
   if hist['filingFrom']<'2026-10-09' and hist['filingTo']>'2016-01-01':
    path=f'data/raw/sec/{ticker}_{hist["name"]}';download('https://data.sec.gov/submissions/'+hist['name'],path);blocks.append(json.loads((ROOT/path).read_text()))
  selected={}
  for block in blocks:
   for i,form in enumerate(block['form']):
    end=block.get('reportDate',['']*len(block['form']))[i];filed=block['filingDate'][i]
    if form not in ['10-K','10-Q'] or filed>CUTOFF or end<'2016-12-31':continue
    if form=='10-K' and end[:4]>'2025':continue
    if form=='10-Q' and end not in ['2025-06-30','2026-06-30']:continue
    key=(form,end)
    if key in selected and selected[key]['filingDate']>filed:continue
    accn=block['accessionNumber'][i];doc=block['primaryDocument'][i];url=f'https://www.sec.gov/Archives/edgar/data/{cik}/{accn.replace("-","")}/{doc}'
    selected[key]={'ticker':ticker,'cik':cik,'form':form,'reportDate':end,'filingDate':filed,'accn':accn,'url':url,'path':f'data/raw/filings/{ticker}_{accn}_{Path(doc).name}'}
  items.extend(selected.values())
 def one(m):
  try:return {**m,**download(m['url'],m['path'])}
  except Exception as e:return {**m,'error':str(e)}
 with ThreadPoolExecutor(max_workers=3) as pool:
  result=[]
  for m in pool.map(one,items):
   result.append(m)
   (ROOT/'data/processed/filing_manifest.json').write_text(json.dumps(result,indent=2))
   print('Filing',m['ticker'],m['reportDate'],m.get('bytes',m.get('error')),flush=True)
 return result
def main(mode):
 (ROOT/'data/processed').mkdir(parents=True,exist_ok=True)
 if mode!='filings':
  results=[]
  with ThreadPoolExecutor(max_workers=3) as pool:
   for b,docs in pool.map(initial,BANKS):results.append((b,docs));print('Data',b['ticker'],b['asof_close'],flush=True)
  (ROOT/'data/processed/banks.json').write_text(json.dumps([x[0] for x in results],indent=2))
  (ROOT/'data/processed/source_manifest.json').write_text(json.dumps([d for _,docs in results for d in docs],indent=2))
 if mode!='initial':filings()
if __name__=='__main__':
 parser=argparse.ArgumentParser();parser.add_argument('--mode',choices=['initial','filings','all'],default='all');main(parser.parse_args().mode)
