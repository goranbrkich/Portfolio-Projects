"""Fetch native issuer PDFs from issuer-linked pages. Original bytes are preserved."""
from fetch_sources import ROOT,download
from concurrent.futures import ThreadPoolExecutor
from urllib.parse import urljoin,unquote
from lxml import html
from pathlib import Path
import json,re
PAGES={'BAC':['https://investor.bankofamerica.com/quarterly-earnings','https://investor.bankofamerica.com/annual-reports-and-proxy-statements'], 'JPM':['https://www.jpmorganchase.com/ir/annual-report','https://www.jpmorganchase.com/ir/quarterly-earnings'], 'C':['https://www.citigroup.com/global/investors/quarterly-earnings/2026','https://www.citigroup.com/global/investors/annual-reports-and-proxy-statements'], 'WFC':['https://www.wellsfargo.com/about/investor-relations/quarterly-earnings/','https://www.wellsfargo.com/about/investor-relations/annual-reports/'], 'GS':['https://www.goldmansachs.com/pressroom/press-releases/2026/2026-07-14-q2-results'], 'MS':['https://www.morganstanley.com/about-us-ir/annual-reports','https://www.morganstanley.com/about-us-ir/earnings-releases'], 'USB':['https://ir.usbank.com/financial-information/financial-results/default.aspx'], 'PNC':['https://investor.pnc.com/financial-information/financial-results'], 'TFC':['https://ir.truist.com/earnings'], 'COF':['https://investor.capitalone.com/financials/annual-reports-and-proxy-statements','https://investor.capitalone.com/financials/quarterly-earnings']}
DIRECT=[('C','H1 2026','https://www.citigroup.com/rcs/citigpa/storage/public/Earnings/Q22026/2026prqtr2rslt.pdf'),('MS','H1 2026','https://www.morganstanley.com/about-us-ir/finsup2q2026/finsup2q2026.pdf'),('MS','H1 2026','https://www.morganstanley.com/about-us-ir/shareholder/2q2026.pdf'),('WFC','H1 2026','https://www.wellsfargo.com/assets/pdf/about/investor-relations/earnings/second-quarter-2026-earnings.pdf'),('JPM','FY2025','https://www.jpmorganchase.com/ir/annual-report/2025/ar-2025.pdf')]
def main():
 candidates=list(DIRECT)
 for ticker,urls in PAGES.items():
  for i,url in enumerate(urls):
   try:
    item=download(url,f'data/raw/reports/{ticker}_index_{i}.html');doc=html.fromstring((ROOT/item['path']).read_bytes())
    for a in doc.xpath('//a[@href]'):
     href=urljoin(url,a.get('href'));text=' '.join(a.text_content().split());name=unquote(href);search=name+' '+text
     if '.pdf' not in href.lower():continue
     period='H1 2026' if re.search(r'(2[qQ][ _+-]?(20)?26|[qQ]2[ _+-]?2026|2026[-_/]q2|second-quarter-2026)',search) else None
     annual=re.search(r'(201[6-9]|202[0-5])',search)
     if period is None and annual and re.search(r'annual|10.?k|ar[-_/]|report',search,re.I):period='FY'+annual[0]
     if period is None:continue
     if period=='H1 2026' and re.search(r'transcript|presentation',search,re.I):continue
     candidates.append((ticker,period,href))
   except Exception as e:print('Index unavailable',ticker,str(e),flush=True)
 unique=[];seen=set()
 for t,p,u in candidates:
  if u not in seen:seen.add(u);unique.append((t,p,u))
 def one(item):
  t,p,u=item;stem=re.sub(r'[^A-Za-z0-9_.-]+','_',unquote(Path(u.split('?')[0]).name))
  path=f'data/raw/reports/{t}_{p.replace(" ","")}_{stem}'
  try:
   m=download(u,path)
   if not (ROOT/path).read_bytes().startswith(b'%PDF'):raise ValueError('Not a native PDF')
   return {'ticker':t,'period':p,**m}
  except Exception as e:print('PDF unavailable',t,p,str(e),flush=True);return None
 result=[]
 with ThreadPoolExecutor(max_workers=3) as pool:
  for m in pool.map(one,unique):
   if m:result.append(m);print('PDF',m['ticker'],m['period'],m['bytes'],flush=True)
   (ROOT/'data/processed/report_manifest.json').write_text(json.dumps(result,indent=2))
 print('Native PDFs retained',len(result),flush=True)
if __name__=='__main__':main()
