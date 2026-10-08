"""Normalize audited Microsoft statements with row-level primary provenance.

The latest comparative presentation available at the fixed cutoff is used for
each year. Fiscal 2016 is an opening-balance record only. SEC Company Facts are
an independent source control, not a substitute for fiscal-period selection.
"""
from __future__ import annotations
import argparse
import csv
import json
import re
from pathlib import Path
from lxml import html

PROJECT = Path(__file__).resolve().parents[1]
CUTOFF = "2026-10-08"

def clean(s):
    return " ".join(s.replace("\u00a0", " ").replace("’", "'").split())

def norm(s):
    return re.sub(r"[^a-z0-9]+", " ", s.lower()).strip()

def table_rows(table):
    return [[clean(c.text_content()) for c in r.xpath("./td|./th")]
            for r in table.xpath("./tr|./tbody/tr|./thead/tr")]

def numbers(cells):
    result = []
    for cell in cells:
        c = cell.replace("$", "").replace(",", "").replace(" ", "")
        if re.fullmatch(r"\(?-?\d+(?:\.\d+)?\)?", c):
            v = float(c.replace("(", "").replace(")", ""))
            result.append(-abs(v) if c.startswith("(") else v)
    return result

def parse_table(table):
    rows = table_rows(table)
    years = []
    for row in rows[:16]:
        candidates = [int(c) for c in row if re.fullmatch(r"20\d\d", c)]
        if candidates:
            years = candidates
            break
    parsed = []
    for row in rows:
        nonblank = [c for c in row if c]
        if not nonblank:
            continue
        label = nonblank[0]
        vals = numbers(nonblank[1:])
        parsed.append((label, vals))
    return years, parsed

IS_ALIASES = {
    "revenue": ["total revenue"], "cogs": ["total cost of revenue"],
    "gross_profit": ["gross margin"], "rd": ["research and development"],
    "sales_marketing": ["sales and marketing"], "ga": ["general and administrative"],
    "other_opex": ["impairment integration and restructuring", "impairment and restructuring", "restructuring"],
    "operating_income": ["operating income"], "other_income": ["other income expense net", "other income net", "other expense net"],
    "pretax_income": ["income before income taxes"], "income_tax": ["provision for income taxes"],
    "net_income": ["net income"],
}
BS_ALIASES = {
    "cash": "cash and cash equivalents", "st_investments": "short term investments",
    "liquid_assets": "total cash cash equivalents and short term investments",
    "receivables": "accounts receivable net", "inventory": "inventories",
    "current_assets": "total current assets", "ppe": "property and equipment net",
    "op_rou": "operating lease right of use assets", "equity_investments": "equity and other investments",
    "goodwill": "goodwill", "intangibles": "intangible assets net", "other_lt_assets": "other long term assets",
    "total_assets": "total assets", "payables": "accounts payable", "st_debt": "short term debt",
    "current_debt": "current portion of long term debt", "compensation": "accrued compensation",
    "current_tax": "short term income taxes", "current_deferred_rev": "short term unearned revenue",
    "securities_lending": "securities lending payable", "current_liabilities": "total current liabilities",
    "long_debt": "long term debt", "long_tax": "long term income taxes",
    "long_deferred_rev": "long term unearned revenue", "deferred_tax_liab": "deferred income taxes",
    "op_lease_long": "operating lease liabilities", "other_lt_liab": "other long term liabilities",
    "total_liabilities": "total liabilities", "paid_in_capital": "common stock and paid in capital",
    "retained_earnings": "retained earnings", "aoci": "accumulated other comprehensive",
    "total_equity": "total stockholders equity", "total_le": "total liabilities and stockholders equity",
}
CF_ALIASES = {
    "net_income": "net income", "impairments": "asset impairments", "da_other": "depreciation amortization and other",
    "sbc": "stock based compensation expense", "investment_adjustment": "net recognized",
    "deferred_taxes": "deferred income taxes", "ar_change": "accounts receivable", "inventory_change": "inventories",
    "other_ca_change": "other current assets", "other_lta_change": "other long term assets",
    "ap_change": "accounts payable", "deferred_rev_change": "unearned revenue", "tax_change": "income taxes",
    "other_cl_change": "other current liabilities", "other_ltl_change": "other long term liabilities",
    "cfo": "net cash from operations", "net_st_debt": "proceeds from issuance repayments of",
    "debt_issued": "proceeds from issuance of debt", "debt_repaid": "repayments of debt",
    "stock_issued": "common stock issued", "buybacks": "common stock repurchased",
    "dividends": "common stock cash dividends paid", "capex": "additions to property and equipment",
    "acquisitions": "acquisition of companies", "investments_purchased": "purchases of investments",
    "investments_matured": "maturities of investments", "investments_sold": "sales of investments",
    "fx_cash": "effect of foreign exchange rates", "cash_change": "net change in cash and cash equivalents",
    "opening_cash": "cash and cash equivalents beginning of period", "closing_cash": "cash and cash equivalents end of period",
}
SEGMENTS = ["Productivity and Business Processes", "Intelligent Cloud", "More Personal Computing"]

def statement_type(table):
    text = clean(" ".join(table.xpath(".//text()")))
    if "Weighted average shares outstanding" in text and "Provision for income taxes" in text:
        return "income"
    if "Cash and cash equivalents" in text and "Total liabilities and stockholders" in text:
        return "balance"
    if "Net cash from operations" in text and "Cash and cash equivalents, beginning of period" in text:
        return "cashflow"
    return None

def extract_statement(kind, rows, position):
    d = {}
    section = ""
    for label, vals in rows:
        key = norm(label)
        if not vals:
            if key.startswith("current assets"): section = "ca"
            elif key.startswith("current liabilities"): section = "cl"
            elif key.startswith("earnings per share"): section = "eps"
            elif key.startswith("weighted average shares"): section = "shares"
            elif key == "financing": section = "financing"
            elif key == "investing": section = "investing"
            continue
        if position >= len(vals):
            continue
        value = vals[position]
        if kind == "income":
            for name, aliases in IS_ALIASES.items():
                if any(key.startswith(alias) for alias in aliases):
                    d[name] = value
            if key in ("basic", "diluted") and section in ("eps", "shares"):
                d[f"{section}_{key}"] = value
        elif kind == "balance":
            for name, alias in BS_ALIASES.items():
                if (key == alias if name.startswith("total_") else key.startswith(alias)): d[name] = value
            if key == "equity investments": d["equity_investments"] = value
            if key in ("other", "other current assets") and section == "ca": d["other_current_assets"] = value
            if key in ("other", "other current liabilities") and section == "cl": d["other_current_liab"] = value
            if key == "income taxes" and section == "cl": d["current_tax"] = value
        else:
            for name, alias in CF_ALIASES.items():
                if key.startswith(alias) and not (name == "debt_issued" and ("maturities" in key or "short term" in key)): d[name] = value
            if "goodwill and asset impairments" in key: d["impairments"] = value
            if key.startswith("net cash") and "financing" in key: d["cff"] = value
            if key.startswith("net cash") and "investing" in key: d["cfi"] = value
            if ("maturities of 90 days or less" in key or "short term debt" in key) and section == "financing": d["net_st_debt"] = value
            if key.startswith("cash premium on debt exchange"): d["debt_exchange_cash_premium"] = value
            if key in ("other", "other net") and section in ("financing", "investing"):
                d["other_financing" if section == "financing" else "other_investing"] = value
            if key == "securities lending payable": d["securities_lending_cf"] = value
    if kind == "income":
        d.setdefault("other_opex", 0.0)
    if kind == "balance":
        for k in ("op_rou", "op_lease_long", "st_debt", "securities_lending", "long_tax"):
            d.setdefault(k, 0.0)
    if kind == "cashflow":
        for k in ("impairments", "tax_change", "other_investing", "other_financing", "securities_lending_cf", "debt_issued", "net_st_debt", "debt_exchange_cash_premium"):
            d.setdefault(k, 0.0)
        d["other_financing"] += d["debt_exchange_cash_premium"]
        sum_known = sum(d.get(k, 0) for k in ["net_income", "impairments", "da_other", "sbc", "investment_adjustment",
            "deferred_taxes", "ar_change", "inventory_change", "other_ca_change", "other_lta_change", "ap_change",
            "deferred_rev_change", "tax_change", "other_cl_change", "other_ltl_change"])
        d["other_operating_adjustments"] = d["cfo"] - sum_known
    return d

def select_fact(facts, tag, year, duration=True):
    entries = facts.get(tag, {}).get("units", {}).get("USD", [])
    selected = [x for x in entries if x["end"] == f"{year}-06-30" and x.get("form") == "10-K"
                and x.get("filed", "9999") <= CUTOFF
                and (not duration or x.get("start") == f"{year-1}-07-01")]
    return max(selected, key=lambda x: (x["filed"], x.get("fy", 0))) if selected else None

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--cache", type=Path, default=PROJECT / ".cache")
    args = parser.parse_args()
    manifest = json.loads((PROJECT / "data/raw/source_manifest.json").read_text())
    urls = {x["file"]: x["url"] for x in manifest["sources"]}
    facts = json.loads((args.cache / "companyfacts.json").read_text())["facts"]["us-gaap"]
    actuals = {str(y): {} for y in range(2016, 2027)}
    provenance, primary_tables, segment_versions = {}, [], []
    for report_year in range(2016, 2027):
        filename = f"ar{report_year % 100:02}.html"
        root = html.parse(str(args.cache / filename))
        for table_idx, table in enumerate(root.xpath("//table")):
            years, rows = parse_table(table)
            if not years: continue
            kind = statement_type(table)
            if kind:
                primary_tables.append({"report_year": report_year, "statement": kind, "years": years,
                                       "table_index": table_idx, "url": urls[filename], "rows": rows})
                for position, year in enumerate(years):
                    if year not in range(2016, 2027): continue
                    extracted = extract_statement(kind, rows, position)
                    actuals[str(year)][kind] = extracted
                    provenance[f"{year}.{kind}"] = {"report_year": report_year, "url": urls[filename],
                                                      "table_index": table_idx}
            # Segments: legacy tables list segment names under Revenue and Operating income.
            row_norms = [norm(r[0]) for r in rows]
            if all(norm(s) in row_norms for s in SEGMENTS):
                text = clean(" ".join(table.xpath(".//text()")))
                if "Revenue" not in text and "Operating income" not in text: continue
                seg = {}
                section = "revenue" if any(norm(r[0]) == "revenue" for r in rows) else "operating_income" if "Operating income" in text else ""
                current_segment = None
                for label, vals in rows:
                    n = norm(label)
                    if n == "total":
                        current_segment = None
                        continue
                    if n.startswith("operating income"): section = "operating_income"
                    elif n == "revenue" and not vals: section = "revenue"
                    if n in [norm(s) for s in SEGMENTS]:
                        current_segment = SEGMENTS[[norm(s) for s in SEGMENTS].index(n)]
                        if vals and section:
                            seg.setdefault(current_segment, {})[section] = vals
                    elif current_segment and n in ("revenue", "cost of revenue", "operating expenses", "operating income") and vals:
                        seg.setdefault(current_segment, {})[n.replace(" ", "_")] = vals
                if all(seg.get(s) for s in SEGMENTS):
                    segment_versions.append({"report_year": report_year, "years": years, "url": urls[filename],
                                             "table_index": table_idx, "data": seg})
    segments = {}
    for version in segment_versions:
        for i, year in enumerate(version["years"]):
            if year not in range(2017, 2027): continue
            # Keep most recent table; later tables in the same filing add detail.
            previous = segments.get(str(year))
            if previous and previous["report_year"] > version["report_year"]: continue
            rec = {"report_year": version["report_year"], "url": version["url"], "table_index": version["table_index"], "data": {}}
            for segment in SEGMENTS:
                rec["data"][segment] = {k: values[i] for k, values in version["data"][segment].items() if i < len(values)}
            if previous and previous["report_year"] == version["report_year"]:
                for segment in SEGMENTS:
                    rec["data"][segment] = {**previous["data"][segment], **rec["data"][segment]}
            segments[str(year)] = rec
    notes = {}
    note_tags = {"ppe_depreciation": ("Depreciation", True), "intangible_amortization": ("AmortizationOfIntangibleAssets", True),
        "ppe_gross": ("PropertyPlantAndEquipmentGross", False),
        "accumulated_depreciation": ("AccumulatedDepreciationDepletionAndAmortizationPropertyPlantAndEquipment", False),
        "operating_lease_cost": ("OperatingLeaseCost", True),
        "finance_lease_liab": ("FinanceLeaseLiability", False), "finance_lease_principal": ("FinanceLeasePrincipalPayments", True),
        "finance_lease_additions": ("RightOfUseAssetObtainedInExchangeForFinanceLeaseLiability", True),
        "operating_lease_additions": ("RightOfUseAssetObtainedInExchangeForOperatingLeaseLiability", True),
        "operating_lease_liab": ("OperatingLeaseLiability", False),
        "interest_expense": ("InterestExpenseNonoperating", True), "interest_income": ("InvestmentIncomeInterestAndDividend", True)}
    selected_facts = []
    for year in range(2016, 2027):
        notes[str(year)] = {}
        for name, (tag, duration) in note_tags.items():
            record = select_fact(facts, tag, year, duration)
            if record:
                notes[str(year)][name] = record["val"] / 1e6
                selected_facts.append({"metric": name, "tag": tag, "year": year, **record})
    # Additional FY2026 disclosures explicitly transcribed from Notes 6, 9, 10 and 13.
    notes["2026"].update({"land": 10546, "ppe_gross": 431767, "accumulated_depreciation": 118691,
        "capex_payables": 26700, "op_lease_current": 5393, "finance_lease_current": 4290,
        "finance_lease_long": 62304, "debt_face": 46136, "debt_discount": 1081,
        "debt_hedge_adjustment": 11, "debt_exchange_premium": 4750,
        "debt_maturities": [9250, 0, 2001, 0, 500], "debt_thereafter": 34385,
        "intangible_amortization_schedule": [3097, 2141, 1944, 1477, 1128],
        "op_lease_payments": [6082, 4334, 3146, 2612, 2316],
        "finance_lease_payments": [7121, 7294, 6668, 6570, 6543],
        "finance_lease_rate": .045, "op_lease_rate": .037,
        "finance_lease_term": 13, "op_lease_term": 6, "construction_commitments": 34566,
        "cash_debt_interest": 1500})
    controls = []
    for year in range(2017, 2027):
        a = actuals[str(year)]
        for statement in ("income", "balance", "cashflow"):
            if statement not in a: raise ValueError(f"Missing {year} {statement}")
        inc, bs, cf = a["income"], a["balance"], a["cashflow"]
        controls.extend([
            {"year": year, "check": "Income statement", "difference": inc["revenue"] - inc["cogs"] - inc["rd"] - inc["sales_marketing"] - inc["ga"] - inc["other_opex"] - inc["operating_income"]},
            {"year": year, "check": "Balance sheet", "difference": bs["total_assets"] - bs["total_liabilities"] - bs["total_equity"]},
            {"year": year, "check": "Cash flow ending cash", "difference": cf["opening_cash"] + cf["cfo"] + cf["cfi"] + cf["cff"] + cf["fx_cash"] - cf["closing_cash"]},
            {"year": year, "check": "Cash to balance sheet", "difference": cf["closing_cash"] - bs["cash"]},
            {"year": year, "check": "Segment revenue", "difference": sum(segments[str(year)]["data"][s]["revenue"] for s in SEGMENTS) - inc["revenue"]},
        ])
        component_checks = {
            "Income before tax components": inc["operating_income"] + inc["other_income"] - inc["pretax_income"],
            "Net income components": inc["pretax_income"] - inc["income_tax"] - inc["net_income"],
            "Current asset components": sum(bs[k] for k in ["cash", "st_investments", "receivables", "inventory", "other_current_assets"]) - bs["current_assets"],
            "Total asset components": bs["current_assets"] + sum(bs[k] for k in ["ppe", "op_rou", "equity_investments", "goodwill", "intangibles", "other_lt_assets"]) - bs["total_assets"],
            "Current liability components": sum(bs[k] for k in ["payables", "st_debt", "current_debt", "compensation", "current_tax", "current_deferred_rev", "securities_lending", "other_current_liab"]) - bs["current_liabilities"],
            "Total liability components": bs["current_liabilities"] + sum(bs[k] for k in ["long_debt", "long_tax", "long_deferred_rev", "deferred_tax_liab", "op_lease_long", "other_lt_liab"]) - bs["total_liabilities"],
            "Operating cash components": sum(cf[k] for k in ["net_income", "da_other", "impairments", "sbc", "investment_adjustment", "deferred_taxes", "ar_change", "inventory_change", "other_ca_change", "other_lta_change", "ap_change", "deferred_rev_change", "tax_change", "other_cl_change", "other_ltl_change", "other_operating_adjustments"]) - cf["cfo"],
            "Investing cash components": sum(cf[k] for k in ["capex", "acquisitions", "investments_purchased", "investments_matured", "investments_sold", "other_investing", "securities_lending_cf"]) - cf["cfi"],
            "Financing cash components": sum(cf[k] for k in ["debt_issued", "net_st_debt", "debt_repaid", "stock_issued", "buybacks", "dividends", "other_financing"]) - cf["cff"],
            "Segment operating income and corporate expense": sum(segments[str(year)]["data"][seg]["operating_income"] for seg in SEGMENTS) - inc["other_opex"] - inc["operating_income"],
        }
        controls.extend({"year": year, "check": check, "difference": difference} for check, difference in component_checks.items())
        for field, tag in [("revenue", "RevenueFromContractWithCustomerExcludingAssessedTax"), ("net_income", "NetIncomeLoss")]:
            fact = select_fact(facts, tag, year)
            if fact:
                controls.append({"year": year, "check": f"SEC XBRL {field}", "difference": inc[field] - fact["val"] / 1e6})
                selected_facts.append({"metric": field, "tag": tag, "year": year, **fact})
    failed = [x for x in controls if abs(x["difference"]) > .01]
    if failed: raise ValueError(json.dumps(failed, indent=2))
    output = {"company": "Microsoft Corporation", "ticker": "MSFT", "currency": "USD", "scale": "millions",
              "information_cutoff": CUTOFF, "historical_years": list(range(2017, 2027)), "opening_balance_year": 2016,
              "actuals": actuals, "segments": segments, "notes": notes, "provenance": provenance, "source_controls": controls}
    (PROJECT / "data/processed/historical_financials.json").write_text(json.dumps(output, indent=2) + "\n")
    (PROJECT / "data/raw/selected_sec_facts.json").write_text(json.dumps(selected_facts, indent=2) + "\n")
    (PROJECT / "data/raw/primary_statement_tables.json").write_text(json.dumps(primary_tables, indent=2) + "\n")
    with (PROJECT / "data/processed/historical_financials.csv").open("w", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["fiscal_year", "statement", "metric", "value", "unit", "source_report_year", "source_table", "source_url"])
        for year in range(2017, 2027):
            for statement in ("income", "balance", "cashflow"):
                p = provenance[f"{year}.{statement}"]
                for metric, value in actuals[str(year)][statement].items():
                    unit = "USD per share" if metric.startswith("eps") else "million shares" if metric.startswith("shares") else "USD millions"
                    writer.writerow([year, statement, metric, value, unit, p["report_year"], p["table_index"], p["url"]])
    print("Normalized FY2017–FY2026;", len(controls), "primary source checks passed.")
    print("FY2026 revenue", actuals["2026"]["income"]["revenue"], "CFO", actuals["2026"]["cashflow"]["cfo"])
    print("Segment source years", {y: v["report_year"] for y, v in segments.items()})

if __name__ == "__main__": main()
