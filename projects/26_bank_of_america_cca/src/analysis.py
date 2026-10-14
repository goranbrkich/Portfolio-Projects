"""Frozen bank CCA, FY2026 earnings, HC3 regression and market risk.
Run: python src/analysis.py --cohort 1
Optional saved Excel edits: python src/analysis.py --from-excel deliverables/Bank_of_America_CCA_Model.xlsx
Requires numpy scipy matplotlib. No credentials or live downloads.
"""
from pathlib import Path
import csv,json,math,argparse,zipfile,xml.etree.ElementTree as ET
import numpy as np
from scipy import stats
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
ROOT=Path(__file__).resolve().parents[1]
SEED=20261008
def raw(r,k):
 v=r['metrics'].get(k);return float(v['value']) if v and v.get('value') is not None else None
def div(a,b):return a/b if a is not None and b is not None and b!=0 else None
def add(*xs):return sum(xs) if all(x is not None for x in xs) else None
def sub(a,b):return a-b if a is not None and b is not None else None
def revenue(r):
 x=add(raw(r,'nii'),raw(r,'noninterest_income'));return x if x is not None else raw(r,'revenue_reported')
def provision(r):
 x=raw(r,'provision');return x if x is not None else sub(sub(revenue(r),raw(r,'opex')),raw(r,'pretax'))
def claims(r):
 x=sub(raw(r,'net_income'),raw(r,'common_income'));return x if x is not None else raw(r,'preferred_div')
def deposits(r):
 x=raw(r,'deposits');components=r.get('deposit_components',{})
 return x if x is not None else add(*(components[k]['value'] if k in components else None for k in ['noninterest','interest']))
def basis(r):
 eq=raw(r,'equity');pref=raw(r,'preferred');gw=raw(r,'goodwill');ia=raw(r,'intangibles');totalia=raw(r,'intangibles_total');shares=raw(r,'shares_end')
 if shares is None:shares=sub(raw(r,'shares_issued'),raw(r,'shares_treasury'))
 common=raw(r,'common_equity_reported');rbv=raw(r,'bvps_reported');rtbv=raw(r,'tbvps_reported');rtce=raw(r,'tce_reported')
 if common is None and pref is not None and pref>=1:common=sub(eq,pref)
 if common is None and rbv is not None and shares is not None:common=rbv*shares
 if totalia is not None and gw is not None:ia=totalia-gw
 standard=sub(sub(common,gw),ia);bv=rbv if rbv is not None else div(common,shares)
 tbv=rtbv if rtbv is not None else div(rtce,shares) if rtce is not None else div(standard,shares)
 tce=rtce if rtce is not None else tbv*shares if tbv is not None and shares is not None else None
 return dict(assets=raw(r,'assets'),common=common,shares=shares,bv=bv,tbv=tbv,tce=tce,standard_tce=standard,tbv_basis='Issuer reported' if rtbv is not None or rtce is not None else 'Straight subtraction' if standard is not None else 'Unavailable')
def ols_hc3(x,y):
 X=np.column_stack([np.ones(len(x)),np.asarray(x)]);y=np.asarray(y,dtype=float)
 if len(y)<=X.shape[1]:raise ValueError('Insufficient residual degrees of freedom')
 bread=np.linalg.inv(X.T@X);coef=bread@X.T@y;resid=y-X@coef;hat=np.sum((X@bread)*X,axis=1)
 if max(hat)>=1-1e-10:raise ValueError('Unit leverage prevents HC3')
 meat=X.T@(X*(resid/(1-hat))[:,None]**2);cov=bread@meat@bread;se=np.sqrt(np.diag(cov));df=len(y)-X.shape[1];critical=stats.t.ppf(.975,df)
 r2=1-np.dot(resid,resid)/np.dot(y-y.mean(),y-y.mean())
 return dict(coef=coef.tolist(),se_hc3=se.tolist(),cov_hc3=cov.tolist(),ci95=np.column_stack([coef-critical*se,coef+critical*se]).tolist(),pvalue=(2*stats.t.sf(np.abs(coef/se),df)).tolist(),n=len(y),df=df,r2=float(r2),residuals=resid.tolist(),leverage=hat.tolist(),critical=float(critical))
def read_excel(path):
 ns={'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'};rel='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
 with zipfile.ZipFile(path) as z:
  strings=[]
  if 'xl/sharedStrings.xml' in z.namelist():strings=[''.join(x.itertext()) for x in ET.fromstring(z.read('xl/sharedStrings.xml'))]
  relationships={r.attrib['Id']:r.attrib['Target'] for r in ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))};result={}
  for s in ET.fromstring(z.read('xl/workbook.xml')).findall('s:sheets/s:sheet',ns):
   target=relationships[s.attrib[rel]];target=target.lstrip('/') if target.startswith('/') else 'xl/'+target;values={}
   for c in ET.fromstring(z.read(target)).findall('.//s:sheetData/s:row/s:c',ns):
    v=c.find('s:v',ns);typ=c.get('t');value=v.text if v is not None else None
    if typ=='s' and value is not None:value=strings[int(value)]
    elif typ=='inlineStr':value=''.join(c.find('s:is',ns).itertext())
    elif value is not None and typ not in ['str','e']:
     try:value=float(value)
     except ValueError:pass
    values[c.get('r')]=value
   result[s.get('name')]=values
 return result
def main(cohort=1,excel=None):
 banks=json.loads((ROOT/'data/processed/banks.json').read_text());records=json.loads((ROOT/'data/processed/financial_inputs.json').read_text());drivers={b['ticker']:[.03,.02,0,.23,-.01] for b in banks};weights=[.25,.25,.5]
 if excel:
  cells=read_excel(excel);a=cells['Assumptions'];cohort=int(a['E6']);weights=[a[f'E{r}'] for r in [23,24,25]];row=7
  for r in records:
   for k,v in r['metrics'].items():
    x=cells['Historical Inputs'].get('F'+str(row));row+=1
    if v is not None:v['value']=x if isinstance(x,(float,int)) else None
    elif isinstance(x,(float,int)):r['metrics'][k]={'value':x}
    if k=='deposits' and r.get('deposit_components'):
     for key,col in [('noninterest','T'),('interest','U')]:
      if key in r['deposit_components']:r['deposit_components'][key]['value']=cells['Historical Inputs'].get(col+str(row-1))
  for i,b in enumerate(banks):b['asof_close']=cells['Market Inputs']['D'+str(i+7)];drivers[b['ticker']]=[a[c+str(i+9)] for c in 'EFGHI']
 lookup={(r['ticker'],r['period']):r for r in records};current=[];forecast=[];stress=[]
 for bank in banks:
  t=bank['ticker'];fy=lookup[t,'FY2025'];old=lookup[t,'H1 2025'];now=lookup[t,'H1 2026'];b=basis(now);bo=basis(old)
  income=add(raw(fy,'common_income'),raw(now,'common_income'),-raw(old,'common_income')) if raw(old,'common_income') is not None else None
  shares=add(raw(fy,'shares_diluted'),raw(now,'shares_diluted')/2,-raw(old,'shares_diluted')/2) if raw(now,'shares_diluted') is not None and raw(old,'shares_diluted') is not None else None
  rev=add(revenue(fy),revenue(now),-revenue(old)) if revenue(old) is not None else None;opex=add(raw(fy,'opex'),raw(now,'opex'),-raw(old,'opex')) if raw(old,'opex') is not None else None
  eps=div(income,shares);tceavg=add(b['tce']/2,bo['tce']/2) if b['tce'] is not None and bo['tce'] is not None else None;ceavg=add(b['common']/2,bo['common']/2) if b['common'] is not None and bo['common'] is not None else None
  price=bank['asof_close'];row={**bank,**b,'common_ltm':income,'shares_ltm':shares,'eps_ltm':eps,'roe_ltm':div(income,ceavg),'rote_ltm':div(income,tceavg),'revenue_ltm':rev,'opex_ltm':opex,'efficiency':div(opex,rev),'cet1':raw(now,'cet1_reported'),'loans':raw(now,'loans_gross'),'deposits':deposits(now),'pe':div(price,eps) if eps and eps>0 else None,'pb':div(price,b['bv']),'ptbv':div(price,b['tbv'])};current.append(row)
  rg,eg,pg,tax,sg=drivers[t];h2rev=sub(revenue(fy),revenue(old));h2opex=sub(raw(fy,'opex'),raw(old,'opex'));h2prov=sub(provision(fy),provision(old))
  h2rev=h2rev*(1+rg) if h2rev is not None else None;h2opex=h2opex*(1+eg) if h2opex is not None else None;h2prov=h2prov*(1+pg) if h2prov is not None else None;h2pretax=sub(sub(h2rev,h2opex),h2prov)
  h2common=sub(h2pretax*(1-tax),claims(now)) if h2pretax is not None else None;fc=add(raw(now,'common_income'),h2common)
  h2shares=sub(2*raw(fy,'shares_diluted'),raw(old,'shares_diluted')) if raw(fy,'shares_diluted') is not None else None;fs=add(raw(now,'shares_diluted')/2,h2shares*(1+sg)/2) if raw(now,'shares_diluted') is not None and h2shares is not None else None
  forecast.append({'ticker':t,'revenue':add(revenue(now),h2rev),'opex':add(raw(now,'opex'),h2opex),'provision_or_residual':add(provision(now),h2prov),'h2_pretax':h2pretax,'common_income':fc,'diluted_shares':fs,'eps':div(fc,fs),'assumptions':drivers[t]})
  loans=row['loans'];stress.append({'ticker':t,'common_income_ltm':income,'loans':loans,'tax':tax,**{f'income_{int(bp*10000)}bp':income-loans*bp*(1-tax) if income is not None and loans is not None else None for bp in [.005,.01,.02]}})
 peers=[r for r in current if r['ticker']!='BAC' and (cohort==1 or r['ticker'] not in ['GS','MS'])];summaries={}
 for k in ['pe','pb','ptbv','roe_ltm','rote_ltm','efficiency','cet1']:
  v=np.array([r[k] for r in peers if r[k] is not None and math.isfinite(r[k])]);summaries[k]={name:float(fn(v)) if len(v) else None for name,fn in [('median',np.median),('mean',np.mean),('q25',lambda a:np.percentile(a,25)),('q75',lambda a:np.percentile(a,75)),('min',np.min),('max',np.max)]};summaries[k]['n']=len(v)
 target=current[0];implied={k:summaries[k]['median']*target[d] if summaries[k]['median'] is not None and target[d] is not None else None for k,d in [('pe','eps_ltm'),('pb','bv'),('ptbv','tbv')]};weighted=sum(implied[k]*w for k,w in zip(['pe','pb','ptbv'],weights)) if all(implied[k] is not None for k in implied) else None
 sample=[r for r in peers if r['rote_ltm'] is not None and r['ptbv'] is not None];reg=ols_hc3([r['rote_ltm'] for r in sample],[r['ptbv'] for r in sample]) if len(sample)>2 else None
 if reg:
  x0=np.array([1,target['rote_ltm']]);coef=np.asarray(reg['coef']);cov=np.asarray(reg['cov_hc3']);prediction=float(x0@coef);sefit=float(np.sqrt(x0@cov@x0));reg.update({'tickers':[r['ticker'] for r in sample],'target_multiple':prediction,'target_price':prediction*target['tbv'],'target_mean_se':sefit,'target_price_mean_ci95':[(prediction-reg['critical']*sefit)*target['tbv'],(prediction+reg['critical']*sefit)*target['tbv']]})
 multi_sample=[r for r in sample if r['cet1'] is not None];multi=ols_hc3([[r['rote_ltm'],r['cet1'],math.log(r['assets'])] for r in multi_sample],[r['ptbv'] for r in multi_sample]) if len(multi_sample)>4 else None
 rng=np.random.default_rng(SEED);bootstrap=[];leaveout=[]
 if reg:
  xs=np.array([r['rote_ltm'] for r in sample]);ys=np.array([r['ptbv'] for r in sample])
  for _ in range(5000):
   ix=rng.integers(0,len(xs),len(xs));xx=xs[ix];yy=ys[ix]
   if np.var(xx)<1e-12:continue
   slope=np.sum((xx-xx.mean())*(yy-yy.mean()))/np.sum((xx-xx.mean())**2);bootstrap.append((yy.mean()+slope*(target['rote_ltm']-xx.mean()))*target['tbv'])
  for i,r in enumerate(sample):
   keep=np.arange(len(sample))!=i;m=ols_hc3(xs[keep],ys[keep]);leaveout.append({'omitted':r['ticker'],'target_price':(m['coef'][0]+m['coef'][1]*target['rote_ltm'])*target['tbv']})
 risk=[];daily={}
 for b in banks:
  with (ROOT/f"data/processed/{b['ticker']}_prices.csv").open() as f:rows=list(csv.DictReader(f))
  p=np.array([float(x['close']) for x in rows]);ret=p[1:]/p[:-1]-1;daily[b['ticker']]=dict(zip([x['date'] for x in rows][1:],ret));x=ret[-252:];q=np.quantile(x,.05);cum=np.cumprod(1+x);dd=cum/np.maximum.accumulate(np.r_[1,cum])[1:]-1
  risk.append({'ticker':b['ticker'],'observations':len(x),'annual_vol':float(x.std(ddof=1)*np.sqrt(252)),'price_return_1y':float(np.prod(1+x)-1),'var95_daily':float(-q),'cvar95_daily':float(-x[x<=q].mean()),'max_drawdown_1y':float(dd.min())})
 shared=sorted(set.intersection(*(set(d) for d in daily.values())))[-252:];sector=np.mean([[daily[b['ticker']][d] for d in shared] for b in banks if b['ticker']!='BAC'],axis=0);beta=ols_hc3(sector,[daily['BAC'][d] for d in shared]);corr=np.corrcoef([[daily[b['ticker']][d] for d in shared] for b in banks]).tolist()
 sensitivity={'cost_of_equity':[.08+i*.01 for i in range(7)],'sustainable_return':[.09+i*.02 for i in range(7)],'book_growth':[.01+i*.01 for i in range(5)],'fixed_growth':.03,'fixed_return':.15};sensitivity['return_by_cost']=[[(r-.03)/(ke-.03)*target['tbv'] for ke in sensitivity['cost_of_equity']] for r in sensitivity['sustainable_return']];sensitivity['growth_by_cost']=[[ (.15-g)/(ke-g)*target['tbv'] for ke in sensitivity['cost_of_equity']] for g in sensitivity['book_growth']]
 result={'cutoff':'2026-10-08','cohort':cohort,'current':current,'peer_statistics':summaries,'implied_prices':implied,'weights':weights,'weighted_price':weighted,'weighted_price_change':weighted/target['asof_close']-1 if weighted is not None else None,'forecast_fy2026':forecast,'credit_stress':stress,'sensitivity':sensitivity,'regression_hc3':reg,'multivariate_hc3':multi,'bootstrap_price_quantiles':np.percentile(bootstrap,[2.5,50,97.5]).tolist() if bootstrap else [],'leave_one_out':leaveout,'market_risk':risk,'sector_beta_hc3':beta,'price_return_correlation':corr,'bootstrap_seed':SEED,'return_basis':'Price returns; cash dividends excluded.'}
 output='python_results_from_excel.json' if excel else 'python_results.json' if cohort==1 else 'python_results_commercial.json';(ROOT/'data/processed'/output).write_text(json.dumps(result,indent=2,allow_nan=False))
 if cohort==1 and not excel:
  plt.rcParams.update({'font.size':10});fig,axs=plt.subplots(1,2,figsize=(12,4));axs[0].bar([r['ticker'] for r in current],[r['ptbv'] or np.nan for r in current],color=['#35434D']+['#82949D']*9);axs[0].set_ylabel('P/TBV (x)');axs[0].set_title('Market price and tangible book value')
  if reg:
   x=np.array([r['rote_ltm'] for r in sample]);y=np.array([r['ptbv'] for r in sample]);axs[1].scatter(x*100,y,color='#35434D')
   for r in sample:axs[1].annotate(r['ticker'],(r['rote_ltm']*100,r['ptbv']),xytext=(4,4),textcoords='offset points',fontsize=8)
   xx=np.linspace(x.min(),x.max(),80);axs[1].plot(xx*100,reg['coef'][0]+reg['coef'][1]*xx,color='#82949D');axs[1].scatter([target['rote_ltm']*100],[target['ptbv']],marker='D',color='#8B6958',label='BAC');axs[1].legend(frameon=False);axs[1].set_xlabel('LTM tangible equity return proxy (%)');axs[1].set_ylabel('P/TBV (x)')
  for a in axs:a.spines[['top','right']].set_visible(False)
  fig.tight_layout();fig.savefig(ROOT/'docs/valuation_and_returns.png',dpi=180);plt.close(fig)
  fig,ax=plt.subplots(figsize=(10,3.4));ax.bar([r['ticker'] for r in risk],[r['annual_vol']*100 for r in risk],color='#82949D');ax.set_ylabel('Annualized daily price volatility (%)');ax.spines[['top','right']].set_visible(False);fig.tight_layout();fig.savefig(ROOT/'docs/market_risk.png',dpi=180);plt.close(fig)
 print(json.dumps({'banks':len(current),'peers':len(peers),'weighted_price':weighted,'regression_n':len(sample),'regression_price':reg['target_price'] if reg else None,'missing_tbv':[r['ticker'] for r in current if r['tbv'] is None]}));return result
if __name__=='__main__':
 p=argparse.ArgumentParser();p.add_argument('--cohort',type=int,choices=[1,2],default=1);p.add_argument('--from-excel',type=Path);a=p.parse_args();main(a.cohort,a.from_excel)
