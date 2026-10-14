"""Select complete-year and first-half reported facts, retaining source units and accessions."""
from pathlib import Path
from datetime import date
import json,re,html as html_std
from fetch_sources import ROOT,BANKS,CUTOFF
TAGS={
'assets':['Assets'],'liabilities':['Liabilities'],'equity':['StockholdersEquity'],
'preferred':['PreferredStockIncludingAdditionalPaidInCapitalNetOfDiscount','PreferredStockCarryingValue','PreferredStockValue'],'common_equity_reported':['CommonStockholdersEquity'],
'goodwill':['Goodwill'],'intangibles':['IntangibleAssetsNetExcludingGoodwill','FiniteLivedIntangibleAssetsNet'],'intangibles_total':['GoodwillAndIntangibleAssetsNet'],
'shares_end':['CommonStockSharesOutstanding'],'shares_issued':['CommonStockSharesIssued'],'shares_treasury':['TreasuryStockCommonShares','TreasuryStockShares'],
'deposits':['Deposits'],'loans_gross':['FinancingReceivableExcludingAccruedInterestBeforeAllowanceForCreditLoss','LoansAndLeasesReceivableNetOfDeferredIncomeBeforeAllowance','LoansAndLeasesReceivableNetOfDeferredIncome'],
'loans_net':['FinancingReceivableExcludingAccruedInterestAfterAllowanceForCreditLoss','LoansAndLeasesReceivableNetOfDeferredIncomeNetOfAllowance','LoansAndLeasesReceivableNetReportedAmount'],'allowance':['FinancingReceivableAllowanceForCreditLossExcludingAccruedInterest','FinancingReceivableAllowanceForCreditLosses','FinancingReceivableAllowanceForCreditLoss','AllowanceForLoanAndLeaseLosses'],
'nonaccrual':['FinancingReceivableNonaccrualStatusRecordedInvestment','LoansAndLeasesReceivableNonperforming'],'cash':['CashAndCashEquivalentsAtCarryingValue'],'aoci':['AccumulatedOtherComprehensiveIncomeLossNetOfTax'],
'nii':['InterestIncomeExpenseNet'],'noninterest_income':['NoninterestIncome'],'revenue_reported':['RevenuesNetOfInterestExpense','Revenues','RevenueFromContractWithCustomerExcludingAssessedTax'],
'opex':['NoninterestExpense'],'provision':['ProvisionForLoanLeaseAndOtherLosses','ProvisionForCreditLosses','ProvisionForLoanAndLeaseLosses'],
'pretax':['IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest','IncomeLossFromContinuingOperationsBeforeIncomeTaxesMinorityInterestAndIncomeLossFromEquityMethodInvestments'],
'tax':['IncomeTaxExpenseBenefit'],'net_income':['NetIncomeLoss'],'common_income':['NetIncomeLossAvailableToCommonStockholdersBasic'],
'preferred_div':['PreferredStockDividendsAndOtherAdjustments','DividendsPreferredStock'],'eps':['EarningsPerShareDiluted'],'shares_diluted':['WeightedAverageNumberOfDilutedSharesOutstanding'],
'dps':['CommonStockDividendsPerShareDeclared','CommonStockDividendsPerShareCashPaid'],'dividends':['PaymentsOfDividendsCommonStock','PaymentsOfDividends'],'buybacks':['PaymentsForRepurchaseOfCommonStock'],
'chargeoffs':['FinancingReceivableRecordedInvestmentWriteoff','FinancingReceivableChargeOffs'],'recoveries':['FinancingReceivableRecoveries','FinancingReceivableRecovery'],
'rwa':['RiskWeightedAssets'],'cet1_reported':['CommonEquityTierOneCapitalRatio'],'nim_reported':['NetInterestMargin'],
'bvps_reported':['BookValuePerShare'],'tbvps_reported':['TangibleBookValuePerShare'],'tce_reported':['TangibleCommonEquity'],
}
FLOW={'nii','noninterest_income','revenue_reported','opex','provision','pretax','tax','net_income','common_income','preferred_div','eps','shares_diluted','dps','dividends','buybacks','chargeoffs','recoveries'}
LABELS={
'equity':'Shareholders equity attributable to parent','preferred':'Preferred equity carrying amount','common_equity_reported':'Reported common shareholders equity','intangibles':'Other intangible assets excluding goodwill','intangibles_total':'Reported goodwill and intangible assets combined','shares_end':'Period-end common shares outstanding','loans_gross':'Gross loans and leases selected definition','loans_net':'Net loans after allowance selected definition','allowance':'Loan allowance for credit losses','aoci':'Accumulated other comprehensive income','nii':'Net interest income','revenue_reported':'Reported total net revenue','opex':'Noninterest expense','provision':'Provision for credit losses','pretax':'Income before income taxes','net_income':'Net income attributable to parent','common_income':'Income available to common shareholders','preferred_div':'Preferred dividends and other common-income adjustments','eps':'Reported diluted EPS','shares_diluted':'Weighted average diluted shares','dps':'Dividends per common share','dividends':'Cash common dividends paid','buybacks':'Common share repurchases paid','chargeoffs':'Gross loan charge-offs','recoveries':'Loan recoveries','rwa':'Risk-weighted assets','cet1_reported':'Actual standardized CET1 ratio','nim_reported':'Issuer reported net interest margin','bvps_reported':'Reported common book value per share','tbvps_reported':'Reported tangible book value per share','tce_reported':'Reported tangible common equity',
}
PER_SHARE={'eps','dps','bvps_reported','tbvps_reported'}
RATIOS={'cet1_reported','nim_reported'}
def valid(f,start,end,key):
 if f.get('end')!=end or f.get('filed','9999')>CUTOFF or f.get('form') not in ['10-K','10-Q','8-K']:return False
 if key in FLOW:return f.get('start')==start
 return True
def selected(company,key,start,end):
 for tag in TAGS[key]:
  entry=company.get('facts',{}).get('us-gaap',{}).get(tag,{})
  unit='USD/shares' if key in PER_SHARE else 'pure' if key in RATIOS else 'shares' if key in ['shares_end','shares_issued','shares_treasury','shares_diluted'] else 'USD'
  candidates=[f for f in entry.get('units',{}).get(unit,[]) if valid(f,start,end,key)]
  if candidates:
   f=max(candidates,key=lambda x:(x.get('filed',''),x.get('accn','')));v=float(f['val']);scaled=v if unit in ['pure','USD/shares'] else v/1e6
   return {'value':scaled,'original_value':v,'unit':unit,'tag':'us-gaap:'+tag,'label':entry.get('label',LABELS.get(key,key)),'filed':f['filed'],'accn':f['accn'],'source_url':f'https://data.sec.gov/api/xbrl/companyfacts/CIK{int(company["cik"]):010}.json','selection_note':'Latest eligible comparative presentation at the information cutoff.'}
 return None
def extract():
 result=[]
 for t,cik,_,_ in BANKS:
  company=json.loads((ROOT/f'data/raw/sec/{t}_companyfacts.json').read_text())
  for period,start,end in [(f'FY{y}',f'{y}-01-01',f'{y}-12-31') for y in range(2016,2026)]+[('H1 2025','2025-01-01','2025-06-30'),('H1 2026','2026-01-01','2026-06-30')]:
   metrics={k:selected(company,k,start,end) for k in TAGS}
   for f in metrics.values():
    if f:f['source_path']=f'data/raw/sec/{t}_companyfacts.json'
   result.append({'ticker':t,'period':period,'start':start,'end':end,'metrics':metrics})
 (ROOT/'data/processed/financial_inputs.json').write_text(json.dumps(result,indent=2))
 (ROOT/'data/processed/metric_definitions.json').write_text(json.dumps({k:{'label':LABELS.get(k,k.replace('_',' ').title()),'unit':'USD/share' if k in PER_SHARE else '%' if k in RATIOS else 'million' if 'shares' in k else 'USDm'} for k in TAGS},indent=2))
 print('Selected',len(result),'bank-period records',flush=True)
def enrich():
 records=json.loads((ROOT/'data/processed/financial_inputs.json').read_text());manifest=json.loads((ROOT/'data/processed/filing_manifest.json').read_text());byacc={m['accn']:m for m in manifest if not m.get('error')}
 for r in records:
  for f in r['metrics'].values():
   if f and f['accn'] in byacc:
    m=byacc[f['accn']];f.update(source_path=m['path'],source_url=m['url'])
 # Parse inline facts with a bounded regex; no full-document DOM is retained.
 custom={'common_equity_reported':['CommonStockholdersEquity','CommonShareholdersEquity','CommonEquity'], 'tce_reported':['TangibleCommonEquity','TangibleCommonStockholdersEquity','TangibleCommonShareholdersEquity'], 'bvps_reported':['BookValuePerCommonShare','BookValuePerShare','CommonStockBookValuePerShare'], 'tbvps_reported':['TangibleBookValuePerCommonShare','TangibleBookValuePerShare'],'loans_gross':['TotalLoans','LoansAndLeasesNetOfUnearnedIncome'], 'cet1_reported':['CommonEquityTier1CapitalRatioStandardized','CommonEquityTierOneCapitalRatioStandardized','CommonEquityTier1CapitalRatio']}
 bydate={(r['ticker'],r['end']):r for r in records}
 missing_before=sum(f is None for r in records for f in r['metrics'].values());count=0
 for m in manifest:
  if m.get('error'):continue
  # Source anchoring for all years; missing consolidated fact enrichment from exact contexts.
  text=(ROOT/m['path']).read_text(errors='replace');contexts={}
  for cm in re.finditer(r'<(?:[\w-]+:)?context\b([^>]+)>(.*?)</(?:[\w-]+:)?context>',text,re.I|re.S):
   attrs=dict(re.findall(r'([\w:]+)\s*=\s*["\']([^"\']*)["\']',cm[1]));body=cm[2]
   if re.search(r'<(?:[\w-]+:)?(?:segment|scenario)\b',body,re.I):continue
   values={tag:re.findall(r'<(?:[\w-]+:)?'+tag+r'[^>]*>(.*?)<',body,re.I) for tag in ['instant','startDate','endDate']}
   end=(values['instant'] or values['endDate'] or [''])[0];start=(values['startDate'] or [''])[0]
   contexts[attrs.get('id','')]=(start,end)
  candidates=[]
  for fm in re.finditer(r'<ix:nonfraction\b([^>]+)>(.*?)</ix:nonfraction>',text,re.I|re.S):
   attrs={k.lower():v for k,v in re.findall(r'([\w:]+)\s*=\s*["\']([^"\']*)["\']',fm[1])};context=contexts.get(attrs.get('contextref',''))
   if not context:continue
   numeric=html_std.unescape(re.sub('<[^>]+>','',fm[2])).replace(',','').replace('\xa0','').replace(' ','').strip()
   if numeric in ['-','—','–']:numeric='0'
   try:v=float(numeric.strip('()'))*(-1 if attrs.get('sign')=='-' or numeric.startswith('(') else 1)*10**int(attrs.get('scale','0'))
   except (ValueError,OverflowError):continue
   candidates.append((attrs,context,v))
  for attrs,(start,end),v in candidates:
   r=bydate.get((m['ticker'],end));name=attrs.get('name','');tag=name.split(':')[-1]
   if not r:continue
   for k,tags in TAGS.items():
    same_tag=tag.lower() in [x.lower() for x in tags+custom.get(k,[])]
    if not same_tag or k in FLOW and start!=r['start']:continue
    f=r['metrics'].get(k)
    unit='USD/shares' if k in PER_SHARE else 'pure' if k in RATIOS else 'shares' if 'shares' in k else 'USD'
    scaled=v if unit in ['pure','USD/shares'] else v/1e6
    if unit=='pure' and scaled>1:scaled/=100
    # A reported requirement/minimum is never accepted as actual regulatory capital.
    if k=='cet1_reported':continue
    if f is None or m['filingDate']>f.get('filed',''):
     r['metrics'][k]={'value':scaled,'original_value':v,'unit':unit,'tag':name,'label':LABELS.get(k,k),'filed':m['filingDate'],'accn':m['accn'],'source_path':m['path'],'source_url':m['url'],'ixid':attrs.get('id',''),'selection_note':'Exact consolidated inline XBRL context and period.'};count+=1
    elif f.get('accn')==m['accn'] and f.get('tag','').lower()==name.lower() and abs(f['value']-scaled)<max(1e-8,abs(scaled)*1e-9):f['ixid']=attrs.get('id','')
  print('Enriched',m['ticker'],m['reportDate'],flush=True)
  # Keep every successful enrichment independently recoverable.
  (ROOT/'data/processed/financial_inputs.json').write_text(json.dumps(records,indent=2))
 print('Inline selections',count,'missing before',missing_before,'after',sum(f is None for r in records for f in r['metrics'].values()),flush=True)
if __name__=='__main__':
 import sys
 enrich() if '--enrich' in sys.argv else extract()
