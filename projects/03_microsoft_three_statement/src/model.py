"""Independent numeric implementation of the integrated forecast.

Amounts are USD millions; shares are millions. Each Excel schedule implements
the same business relationships with editable native formulas. Assumptions are
research assumptions, not Microsoft guidance or the author's investment thesis.
"""
from __future__ import annotations
import argparse
import copy
import csv
import json
from pathlib import Path

PROJECT = Path(__file__).resolve().parents[1]
YEARS = list(range(2027, 2032))
SEGMENTS = ["Productivity and Business Processes", "Intelligent Cloud", "More Personal Computing"]

def make_assumptions(data):
    a = data["actuals"]["2026"]
    b, i, n = a["balance"], a["income"], data["notes"]["2026"]
    rev = i["revenue"]
    cost_alloc = [(.05, .20, .20), (.75, .05, .60), (.05, .65, .10)]
    op_cost = n.get("operating_lease_cost", 6443)
    cost_rates = [(data["segments"]["2026"]["data"][s]["cost_of_revenue"] - n["ppe_depreciation"]*pa - n["intangible_amortization"]*ia - op_cost*oa) / data["segments"]["2026"]["data"][s]["revenue"] for s,(pa,ia,oa) in zip(SEGMENTS,cost_alloc)]
    shared = {
        "dso": b["receivables"]/rev*365, "inventory_days": b["inventory"]/i["cogs"]*365,
        "dpo": (b["payables"]-n["capex_payables"])/i["cogs"]*365,
        "comp_ratio": b["compensation"]/rev, "deferred_rev_ratio": b["current_deferred_rev"]/rev,
        "other_cl_ratio": (b["other_current_liab"]-n["op_lease_current"]-n["finance_lease_current"])/rev,
        "other_lta_ratio": b["other_lt_assets"]/rev, "long_tax_ratio": b["long_tax"]/rev,
        "long_deferred_rev_ratio": b["long_deferred_rev"]/rev,
        "other_ltl_ratio": (b["other_lt_liab"]-n["finance_lease_long"])/rev,
        "dtl_ratio": b["deferred_tax_liab"]/rev, "min_cash": 25000, "min_st_investments": 20000,
        "debt_coupon": .035, "investment_yield": .035, "capitalized_interest_share": .20,
        "debt_contra_amort_rate": .03, "fin_lease_rate": .045, "op_lease_rate": .037,
        "new_finance_lease_term": 13, "new_op_lease_term": 6, "legacy_op_rou_life": 6,
        "extra_debt": 0, "st_investment_growth": 0, "equity_investment_ratio": .01,
        "stock_issue_ratio": .006, "sbc_ratio": .0375, "buyback_retained_share": .75,
        "dividend_payout": .20, "tax_rate": .20,
    }
    driver_names = {
        "pbp_growth":"Productivity and Business Processes revenue growth", "ic_growth":"Intelligent Cloud revenue growth", "mpc_growth":"More Personal Computing revenue growth",
        "pbp_cost_rate":"PBP cost excluding allocated D&A and operating leases", "ic_cost_rate":"IC cost excluding allocated D&A and operating leases", "mpc_cost_rate":"MPC cost excluding allocated D&A and operating leases",
        "rd_rate":"R&D excluding allocated D&A and operating leases / revenue", "sm_rate":"Sales and marketing / revenue", "ga_rate":"General and administrative / revenue",
        "cash_capex_ratio":"Cash capex including capitalized interest / revenue", "capex_ap_ratio":"Capital expenditure payables / cash capex", "finance_lease_add_ratio":"New finance lease assets / revenue", "op_lease_add_ratio":"New operating lease assets / revenue",
        "legacy_ppe_life":"Legacy depreciable net PPE remaining life in years", "new_ppe_life":"New PPE useful life in years", "dso":"Receivable days on revenue", "inventory_days":"Inventory days on cost of revenue", "dpo":"Operating payable days on cost of revenue",
        "other_ca_ratio":"Other current assets / revenue", "comp_ratio":"Accrued compensation / revenue", "deferred_rev_ratio":"Current unearned revenue / revenue", "other_cl_ratio":"Other current liabilities excluding leases / revenue",
        "other_lta_ratio":"Other long-term assets / revenue", "long_tax_ratio":"Long-term income tax liabilities / revenue", "long_deferred_rev_ratio":"Long-term unearned revenue / revenue", "other_ltl_ratio":"Other long-term liabilities excluding finance leases / revenue", "dtl_ratio":"Deferred tax liabilities / revenue",
        "min_cash":"Minimum cash balance", "min_st_investments":"Minimum short-term investment reserve", "debt_coupon":"Cash coupon rate on opening debt face value", "investment_yield":"Yield on opening cash and short-term investments",
        "capitalized_interest_share":"Cash interest capitalized within the capex budget", "debt_contra_amort_rate":"Annual amortization of debt accounting adjustments", "fin_lease_rate":"Finance lease interest rate", "op_lease_rate":"Operating lease interest rate",
        "new_finance_lease_term":"New finance lease principal term in years", "new_op_lease_term":"New operating lease term in years", "legacy_op_rou_life":"Legacy operating ROU remaining life in years",
        "extra_debt":"Planned new debt face value", "st_investment_growth":"Planned short-term investment balance growth", "equity_investment_ratio":"Cash purchases of equity investments / revenue",
        "stock_issue_ratio":"Cash proceeds from employee share issuance / revenue", "sbc_ratio":"Stock-based compensation / revenue", "buyback_retained_share":"Share of buybacks charged to retained earnings", "dividend_payout":"Dividends as share of positive net income", "tax_rate":"Effective income tax rate",
        "buyback_ratio":"Cash share repurchases / revenue", "diluted_share_growth":"Net change in diluted weighted average shares",
    }
    cases = {}
    for case in ("Base", "Upside", "Downside"):
        cases[case] = {k:[v]*5 for k,v in shared.items()}
    cases["Base"].update({"pbp_growth":[.14,.13,.12,.11,.10],"ic_growth":[.24,.22,.20,.18,.16],"mpc_growth":[.01,.02,.02,.02,.02],
        "cash_capex_ratio":[.32,.28,.25,.22,.20],"capex_ap_ratio":[.20,.18,.16,.14,.12],"finance_lease_add_ratio":[.075,.07,.065,.06,.055],"op_lease_add_ratio":[.015]*5,
        "legacy_ppe_life":[8]*5,"new_ppe_life":[6]*5,"other_ca_ratio":[.135,.125,.12,.115,.11],"buyback_ratio":[.065]*5,"diluted_share_growth":[-.003]*5})
    cases["Upside"].update({"pbp_growth":[.18,.16,.14,.13,.12],"ic_growth":[.32,.28,.25,.22,.20],"mpc_growth":[.04,.04,.03,.03,.03],
        "cash_capex_ratio":[.34,.31,.28,.25,.23],"capex_ap_ratio":[.22,.20,.18,.16,.14],"finance_lease_add_ratio":[.085,.08,.075,.07,.065],"op_lease_add_ratio":[.016]*5,
        "legacy_ppe_life":[9]*5,"new_ppe_life":[6]*5,"other_ca_ratio":[.13,.12,.115,.11,.105],"buyback_ratio":[.07]*5,"diluted_share_growth":[-.005]*5})
    cases["Downside"].update({"pbp_growth":[.08,.07,.06,.06,.05],"ic_growth":[.14,.12,.10,.10,.09],"mpc_growth":[-.05,-.03,0,.01,.01],
        "cash_capex_ratio":[.35,.33,.30,.28,.25],"capex_ap_ratio":[.25,.25,.24,.24,.22],"finance_lease_add_ratio":[.075,.065,.055,.05,.045],"op_lease_add_ratio":[.012]*5,
        "legacy_ppe_life":[7]*5,"new_ppe_life":[5]*5,"other_ca_ratio":[.146,.145,.14,.135,.13],"buyback_ratio":[.045]*5,"diluted_share_growth":[.002]*5})
    rd_base = (i["rd"]-n["ppe_depreciation"]*.15-n["intangible_amortization"]*.10-op_cost*.10)/rev
    for case, adjustment in [("Base",0),("Upside",-.005),("Downside",.01)]:
        for idx,key in enumerate(["pbp_cost_rate","ic_cost_rate","mpc_cost_rate"]):
            cases[case][key]=[max(0,cost_rates[idx]+adjustment-.002*j) for j in range(5)]
        cases[case]["rd_rate"]=[rd_base+(.003 if case=="Downside" else -.001 if case=="Upside" else 0)-.0005*j for j in range(5)]
        cases[case]["sm_rate"]=[.08+(.003 if case=="Downside" else -.003 if case=="Upside" else 0)-.001*j for j in range(5)]
        cases[case]["ga_rate"]=[.024+(.002 if case=="Downside" else -.001 if case=="Upside" else 0)-.0005*j for j in range(5)]
    return {"forecast_years":YEARS,"cases":cases,"driver_names":driver_names,
            "allocation":{"ppe":[.05,.75,.05,.15],"intangibles":[.20,.05,.65,.10],"operating_leases":[.20,.60,.10,.10]},
            "fy2032_debt_maturity_assumption":0,"fy2032_legacy_finance_lease_payment_assumption":6500,
            "fy2032_legacy_operating_lease_payment_assumption":2000,
            "assumption_basis":"Research assumptions dated 8 October 2026. Contractual FY2027–FY2031 payments come from FY2026 notes."}

def forecast(data, assumptions, case="Base", overrides=None):
    drivers = copy.deepcopy(assumptions["cases"][case])
    if overrides:
        for key,values in overrides.items(): drivers[key]=values
    b0=data["actuals"]["2026"]["balance"]; i0=data["actuals"]["2026"]["income"]; n=data["notes"]["2026"]
    previous=copy.deepcopy(b0)
    prev_seg=[data["segments"]["2026"]["data"][s]["revenue"] for s in SEGMENTS]
    legacy_ppe=n["ppe_gross"]-n["accumulated_depreciation"]-n["land"]
    legacy_op_rou=b0["op_rou"]; legacy_op_liab=n["operating_lease_liab"]; legacy_fin_liab=n["finance_lease_liab"]
    new_op_balance=0.; new_fin_balance=0.; cohorts=[]
    debt_face=n["debt_face"]; debt_contra=n["debt_discount"]+n["debt_hedge_adjustment"]+n["debt_exchange_premium"]
    diluted_shares=i0["shares_diluted"]; ppe_gross=n["ppe_gross"]; accum_dep=n["accumulated_depreciation"]; capex_ap=n["capex_payables"]
    output=[]
    for idx,year in enumerate(YEARS):
        p={k:v[idx] for k,v in drivers.items()}
        seg_rev=[v*(1+p[g]) for v,g in zip(prev_seg,["pbp_growth","ic_growth","mpc_growth"])]
        revenue=sum(seg_rev)
        cash_capex=revenue*p["cash_capex_ratio"]
        ending_capex_ap=cash_capex*p["capex_ap_ratio"]
        capex_ap_change=ending_capex_ap-capex_ap
        fin_add=revenue*p["finance_lease_add_ratio"]
        op_add=revenue*p["op_lease_add_ratio"]
        ppe_add=cash_capex+capex_ap_change+fin_add
        legacy_dep=min(legacy_ppe,(b0["ppe"]-n["land"])/p["legacy_ppe_life"])
        legacy_ppe-=legacy_dep
        cohort_dep=[min(c["net"],c["cost"]/c["life"]) for c in cohorts]
        for c,dep in zip(cohorts,cohort_dep): c["net"]-=dep
        current_dep=ppe_add/p["new_ppe_life"]*.5
        cohorts.append({"cost":ppe_add,"life":p["new_ppe_life"],"net":ppe_add-current_dep})
        depreciation=legacy_dep+sum(cohort_dep)+current_dep
        ppe_gross+=ppe_add;accum_dep+=depreciation
        amort=min(previous["intangibles"],n["intangible_amortization_schedule"][idx])
        fin_interest_legacy=legacy_fin_liab*p["fin_lease_rate"]
        fin_principal_legacy=min(legacy_fin_liab,max(0,n["finance_lease_payments"][idx]-fin_interest_legacy))
        fin_interest_new=(new_fin_balance+.5*fin_add)*p["fin_lease_rate"]
        fin_principal_new=min(new_fin_balance+fin_add,new_fin_balance/p["new_finance_lease_term"]+.5*fin_add/p["new_finance_lease_term"])
        end_legacy_fin=legacy_fin_liab-fin_principal_legacy; end_new_fin=new_fin_balance+fin_add-fin_principal_new
        fin_interest=fin_interest_legacy+fin_interest_new; fin_principal=fin_principal_legacy+fin_principal_new
        op_interest=legacy_op_liab*p["op_lease_rate"]
        op_principal=min(legacy_op_liab,max(0,n["op_lease_payments"][idx]-op_interest))
        legacy_rou_amort=min(legacy_op_rou,b0["op_rou"]/p["legacy_op_rou_life"])
        new_op_amort=min(new_op_balance+op_add,new_op_balance/p["new_op_lease_term"]+.5*op_add/p["new_op_lease_term"])
        new_op_interest=(new_op_balance+.5*op_add)*p["op_lease_rate"]
        op_expense=legacy_rou_amort+op_interest+new_op_amort+new_op_interest
        end_legacy_op_rou=legacy_op_rou-legacy_rou_amort; end_legacy_op_liab=legacy_op_liab-op_principal
        end_new_op=new_op_balance+op_add-new_op_amort
        alloc=assumptions["allocation"]
        segment_cost=[r*p[key]+depreciation*alloc["ppe"][j]+amort*alloc["intangibles"][j]+op_expense*alloc["operating_leases"][j]
                      for j,(r,key) in enumerate(zip(seg_rev,["pbp_cost_rate","ic_cost_rate","mpc_cost_rate"]))]
        cogs=sum(segment_cost)
        rd=revenue*p["rd_rate"]+depreciation*alloc["ppe"][3]+amort*alloc["intangibles"][3]+op_expense*alloc["operating_leases"][3]
        sm=revenue*p["sm_rate"]; ga=revenue*p["ga_rate"]
        operating_income=revenue-cogs-rd-sm-ga
        contra_amort=min(debt_contra,debt_contra*p["debt_contra_amort_rate"])
        debt_cash_interest=debt_face*p["debt_coupon"]
        capitalized_interest=(debt_cash_interest+fin_interest)*p["capitalized_interest_share"]
        interest_expense=debt_cash_interest+fin_interest-capitalized_interest+contra_amort
        interest_income=(previous["cash"]+previous["st_investments"])*p["investment_yield"]
        pretax=operating_income+interest_income-interest_expense
        tax=pretax*p["tax_rate"]; net_income=pretax-tax
        dividends=max(0,net_income)*p["dividend_payout"]
        buybacks=revenue*p["buyback_ratio"]; sbc=revenue*p["sbc_ratio"]; stock_issued=revenue*p["stock_issue_ratio"]
        diluted_shares*=1+p["diluted_share_growth"]
        next_debt=n["debt_maturities"][idx+1] if idx<4 else assumptions["fy2032_debt_maturity_assumption"]
        next_fin_pay=n["finance_lease_payments"][idx+1] if idx<4 else assumptions["fy2032_legacy_finance_lease_payment_assumption"]
        next_op_pay=n["op_lease_payments"][idx+1] if idx<4 else assumptions["fy2032_legacy_operating_lease_payment_assumption"]
        fin_total=end_legacy_fin+end_new_fin
        fin_current=min(fin_total,max(0,next_fin_pay-end_legacy_fin*p["fin_lease_rate"])+end_new_fin/p["new_finance_lease_term"])
        op_total=end_legacy_op_liab+end_new_op
        op_current=min(op_total,max(0,next_op_pay-end_legacy_op_liab*p["op_lease_rate"])+end_new_op/p["new_op_lease_term"])
        b={"receivables":revenue*p["dso"]/365,"inventory":cogs*p["inventory_days"]/365,"other_current_assets":revenue*p["other_ca_ratio"],
           "ppe":ppe_gross-accum_dep,"op_rou":end_legacy_op_rou+end_new_op,"equity_investments":previous["equity_investments"]+revenue*p["equity_investment_ratio"],
           "goodwill":previous["goodwill"],"intangibles":previous["intangibles"]-amort,"other_lt_assets":revenue*p["other_lta_ratio"],
           "payables":cogs*p["dpo"]/365+ending_capex_ap,"st_debt":0,"compensation":revenue*p["comp_ratio"],"current_tax":previous["current_tax"],
           "current_deferred_rev":revenue*p["deferred_rev_ratio"],"securities_lending":0,"other_current_liab":revenue*p["other_cl_ratio"]+fin_current+op_current,
           "long_tax":revenue*p["long_tax_ratio"],"long_deferred_rev":revenue*p["long_deferred_rev_ratio"],"deferred_tax_liab":revenue*p["dtl_ratio"],
           "op_lease_long":op_total-op_current,"other_lt_liab":revenue*p["other_ltl_ratio"]+fin_total-fin_current,
           "paid_in_capital":previous["paid_in_capital"]+sbc+stock_issued-buybacks*(1-p["buyback_retained_share"]),
           "retained_earnings":previous["retained_earnings"]+net_income-dividends-buybacks*p["buyback_retained_share"],"aoci":previous["aoci"]}
        previous_op_ap=previous["payables"]-capex_ap
        op_ap=b["payables"]-ending_capex_ap
        cfo_adjustments={"ar_change":previous["receivables"]-b["receivables"],"inventory_change":previous["inventory"]-b["inventory"],
            "other_ca_change":previous["other_current_assets"]-b["other_current_assets"],"other_lta_change":previous["other_lt_assets"]-b["other_lt_assets"],
            "ap_change":op_ap-previous_op_ap,"deferred_rev_change":b["current_deferred_rev"]+b["long_deferred_rev"]-previous["current_deferred_rev"]-previous["long_deferred_rev"],
            "tax_change":b["current_tax"]+b["long_tax"]-previous["current_tax"]-previous["long_tax"],
            "other_cl_change":(b["compensation"]-previous["compensation"])+revenue*p["other_cl_ratio"]-(previous["other_current_liab"]-(n["finance_lease_current"] if idx==0 else output[-1]["fin_current"])-(n["op_lease_current"] if idx==0 else output[-1]["op_current"])),
            "other_ltl_change":revenue*p["other_ltl_ratio"]-(previous["other_lt_liab"]-(n["finance_lease_long"] if idx==0 else output[-1]["fin_total"]-output[-1]["fin_current"])),
            "operating_lease_reconciliation":legacy_rou_amort+new_op_amort-op_principal-new_op_amort}
        deferred_tax_change=b["deferred_tax_liab"]-previous["deferred_tax_liab"]
        cfo=net_income+depreciation+amort+sbc+contra_amort+deferred_tax_change+sum(cfo_adjustments.values())
        st_planned=max(p["min_st_investments"],previous["st_investments"]*(1+p["st_investment_growth"]))
        cfi_before=-cash_capex-(st_planned-previous["st_investments"])-revenue*p["equity_investment_ratio"]
        repayment=min(debt_face,n["debt_maturities"][idx])
        cff_before=p["extra_debt"]-repayment-fin_principal+stock_issued-buybacks-dividends
        cash_before_liquidity=previous["cash"]+cfo+cfi_before+cff_before
        liquidation=min(max(0,st_planned-p["min_st_investments"]),max(0,p["min_cash"]-cash_before_liquidity))
        borrowing=max(0,p["min_cash"]-cash_before_liquidity-liquidation)
        b["cash"]=cash_before_liquidity+liquidation+borrowing;b["st_investments"]=st_planned-liquidation
        ending_face=debt_face-repayment+p["extra_debt"]+borrowing
        ending_contra=debt_contra-contra_amort
        book_debt=ending_face-ending_contra
        b["current_debt"]=min(book_debt,next_debt*(1-23/9250));b["long_debt"]=book_debt-b["current_debt"]
        b["liquid_assets"]=b["cash"]+b["st_investments"]
        b["current_assets"]=sum(b[k] for k in ["cash","st_investments","receivables","inventory","other_current_assets"])
        b["total_assets"]=b["current_assets"]+sum(b[k] for k in ["ppe","op_rou","equity_investments","goodwill","intangibles","other_lt_assets"])
        b["current_liabilities"]=sum(b[k] for k in ["payables","st_debt","current_debt","compensation","current_tax","current_deferred_rev","securities_lending","other_current_liab"])
        b["total_liabilities"]=b["current_liabilities"]+sum(b[k] for k in ["long_debt","long_tax","long_deferred_rev","deferred_tax_liab","op_lease_long","other_lt_liab"])
        b["total_equity"]=b["paid_in_capital"]+b["retained_earnings"]+b["aoci"];b["total_le"]=b["total_liabilities"]+b["total_equity"]
        cfi=cfi_before+liquidation;cff=cff_before+borrowing
        record={"year":year,"case":case,"drivers":p,"balance":b,"revenue":revenue,"segment_revenue":seg_rev,"segment_cost":segment_cost,"cogs":cogs,"rd":rd,"sales_marketing":sm,"ga":ga,"operating_income":operating_income,
                "pretax_income":pretax,"income_tax":tax,"net_income":net_income,"shares_diluted":diluted_shares,"eps_diluted":net_income/diluted_shares,
                "depreciation":depreciation,"legacy_depreciation":legacy_dep,"new_cohort_depreciation":sum(cohort_dep)+current_dep,"cohort_depreciation":cohort_dep+[current_dep],"cohort_net":[c["net"] for c in cohorts],"legacy_ppe":legacy_ppe,
                "ppe_gross":ppe_gross,"accumulated_depreciation":accum_dep,"cash_capex":cash_capex,"capex_payables":ending_capex_ap,"capex_ap_change":capex_ap_change,"ppe_additions":ppe_add,"amortization":amort,
                "fin_additions":fin_add,"fin_interest":fin_interest,"fin_principal":fin_principal,"fin_total":fin_total,"fin_current":fin_current,"fin_legacy":end_legacy_fin,"fin_new":end_new_fin,
                "op_additions":op_add,"op_expense":op_expense,"op_total":op_total,"op_current":op_current,"op_legacy_liab":end_legacy_op_liab,"op_new":end_new_op,"op_legacy_rou":end_legacy_op_rou,"op_rou_amort":legacy_rou_amort+new_op_amort,"op_principal":op_principal+new_op_amort,
                "debt_face":ending_face,"debt_contra":ending_contra,"book_debt":book_debt,"contra_amort":contra_amort,"debt_repayment":repayment,"debt_cash_interest":debt_cash_interest,"capitalized_interest":capitalized_interest,"interest_expense":interest_expense,"interest_income":interest_income,
                "dividends":dividends,"buybacks":buybacks,"sbc":sbc,"stock_issued":stock_issued,"cfo_adjustments":cfo_adjustments,"deferred_tax_change":deferred_tax_change,"cfo":cfo,"cfi":cfi,"cff":cff,"cfi_before":cfi_before,"cff_before":cff_before,
                "st_planned":st_planned,"liquidation":liquidation,"borrowing":borrowing,"cash_before_liquidity":cash_before_liquidity,"fcf":cfo-cash_capex,"fcf_after_finance_leases":cfo-cash_capex-fin_principal}
        output.append(record)
        previous=b;prev_seg=seg_rev;capex_ap=ending_capex_ap;legacy_fin_liab=end_legacy_fin;new_fin_balance=end_new_fin;legacy_op_liab=end_legacy_op_liab;new_op_balance=end_new_op;legacy_op_rou=end_legacy_op_rou;debt_face=ending_face;debt_contra=ending_contra
    return output

def validate(data, assumptions, results):
    checks=[]
    for case, records in results.items():
        opening=data["actuals"]["2026"]["balance"]
        for r in records:
            b=r["balance"]
            items={"Balance sheet":b["total_assets"]-b["total_le"],"Cash rollforward":opening["cash"]+r["cfo"]+r["cfi"]+r["cff"]-b["cash"],
                   "PPE cohorts":r["legacy_ppe"]+data["notes"]["2026"]["land"]+sum(r["cohort_net"])-b["ppe"],"Debt carrying value":r["debt_face"]-r["debt_contra"]-b["current_debt"]-b["long_debt"],
                   "Equity rollforward":opening["total_equity"]+r["net_income"]+r["sbc"]+r["stock_issued"]-r["dividends"]-r["buybacks"]-b["total_equity"],
                   "Segment revenue":sum(r["segment_revenue"])-r["revenue"]}
            for name,difference in items.items():
                checks.append({"case":case,"year":r["year"],"check":name,"difference":difference})
                if abs(difference)>1e-6: raise ValueError(checks[-1])
            if b["cash"]+1e-6<r["drivers"]["min_cash"]: raise ValueError("Cash policy violation")
            if b["st_investments"]+1e-6<r["drivers"]["min_st_investments"]: raise ValueError("Reserve policy violation")
            opening=b
    stress=copy.deepcopy(assumptions["cases"]["Base"]["cash_capex_ratio"]);stress[-1]+=.01
    changed=forecast(data,assumptions,"Base",{"cash_capex_ratio":stress})
    base=results["Base"]
    if changed[-1]["balance"]["ppe"]<=base[-1]["balance"]["ppe"] or changed[-1]["fcf"]>=base[-1]["fcf"]:
        raise ValueError("Later-period capex response failed")
    return {"historical_primary_checks":len(data["source_controls"]),"forecast_checks":checks,"forecast_checks_passed":len(checks),
            "later_period_driver_test":"FY2031 capex +1 percentage point increases net PPE and reduces FCF",
            "all_cases_balanced":True,"maximum_absolute_difference":max(abs(x["difference"]) for x in checks)}

def main():
    parser=argparse.ArgumentParser();parser.add_argument("--case",choices=["Base","Upside","Downside"],default="Base");args=parser.parse_args()
    data=json.loads((PROJECT/"data/processed/historical_financials.json").read_text())
    assumption_path=PROJECT/"data/processed/assumptions.json"
    assumptions=json.loads(assumption_path.read_text()) if assumption_path.exists() else make_assumptions(data)
    assumption_path.write_text(json.dumps(assumptions,indent=2)+"\n")
    results={case:forecast(data,assumptions,case) for case in ["Base","Upside","Downside"]}
    validation=validate(data,assumptions,results)
    (PROJECT/"data/processed/forecast_results.json").write_text(json.dumps(results,indent=2)+"\n")
    (PROJECT/"data/processed/validation_results.json").write_text(json.dumps(validation,indent=2)+"\n")
    with (PROJECT/"data/processed/forecast_summary.csv").open("w", newline="") as handle:
        fields=["case", "year", "revenue", "operating_income", "net_income", "eps_diluted", "cfo", "cash_capex", "fcf", "fcf_after_finance_leases", "depreciation", "book_debt", "fin_total", "op_total", "borrowing"]
        writer=csv.DictWriter(handle,fieldnames=fields);writer.writeheader()
        writer.writerows({key:r[key] for key in fields} for records in results.values() for r in records)
    for r in results[args.case]:print(r["year"],"Revenue",round(r["revenue"],1),"Net income",round(r["net_income"],1),"FCF",round(r["fcf"],1),"Borrowing",round(r["borrowing"],1),"BS difference",round(r["balance"]["total_assets"]-r["balance"]["total_le"],6))
    print("Validated",validation["forecast_checks_passed"],"forecast relationships across all three cases.")

if __name__=="__main__":main()
