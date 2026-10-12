"""Prepare auditable inputs from project 03 without changing its deliverables."""
from __future__ import annotations
import csv
import datetime as dt
import json
import re
import shutil
from pathlib import Path

import fitz
import numpy as np

from valuation import bridge, regression, reverse, value, wacc, monte_carlo

ROOT = Path(__file__).resolve().parents[1]
PRIOR = ROOT.parents[1] / "source_project03/Microsoft_Three_Statement_Project03"

FIELDS = [
    ("income", "revenue", "Revenue"), ("income", "operating_income", "Operating income"),
    ("income", "income_tax", "Income tax expense"), ("income", "pretax_income", "Pretax income"),
    ("income", "shares_basic", "Weighted average basic shares"), ("income", "shares_diluted", "Weighted average diluted shares"),
    ("cashflow", "da_other", "Cash flow depreciation and other adjustments"), ("cashflow", "cfo", "Operating cash flow"),
    ("cashflow", "capex", "Cash purchases of property and equipment"), ("cashflow", "sbc", "Stock based compensation"),
    ("balance", "cash", "Cash and cash equivalents"), ("balance", "st_investments", "Short term investments"),
    ("balance", "equity_investments", "Equity and other investments"), ("balance", "receivables", "Accounts receivable"),
    ("balance", "inventory", "Inventory"), ("balance", "other_current_assets", "Other current assets"),
    ("balance", "payables", "Accounts payable including capex payables"), ("balance", "compensation", "Accrued compensation"),
    ("balance", "current_deferred_rev", "Current unearned revenue"), ("balance", "other_current_liab", "Other current liabilities including lease portions"),
    ("balance", "current_debt", "Current portion of long term debt"), ("balance", "st_debt", "Short term debt"),
    ("balance", "long_debt", "Long term debt carrying amount"), ("balance", "ppe", "Property and equipment net"),
    ("notes", "ppe_depreciation", "Property and equipment depreciation"), ("notes", "intangible_amortization", "Intangible amortization"),
    ("notes", "finance_lease_liab", "Finance lease liabilities"), ("notes", "finance_lease_additions", "Finance lease asset additions"),
    ("notes", "finance_lease_principal", "Finance lease principal paid"), ("notes", "capex_payables", "Capital expenditure payables"),
    ("notes", "op_lease_current", "Current operating lease liability"), ("notes", "finance_lease_current", "Current finance lease liability"),
    ("notes", "operating_lease_liab", "Operating lease liabilities"), ("notes", "operating_lease_cost", "Operating lease expense"),
    ("notes", "debt_face", "Debt principal outstanding"), ("notes", "interest_expense", "Interest expense"),
    ("notes", "land", "Land cost"),
    ("segments", "pbp_revenue", "Productivity and Business Processes revenue"),
    ("segments", "ic_revenue", "Intelligent Cloud revenue"),
    ("segments", "mpc_revenue", "More Personal Computing revenue"),
]


def nwc(balance, note):
    return (balance["receivables"] + balance["inventory"] + balance["other_current_assets"]
            - (balance["payables"] - note.get("capex_payables", 0)) - balance["compensation"]
            - balance["current_deferred_rev"]
            - (balance["other_current_liab"] - note.get("op_lease_current", 0) - note.get("finance_lease_current", 0)))


def main():
    for name in ["data/processed", "data/raw/reports/originals", "data/raw/reports/highlighted", "data/raw/source_snapshots", "docs", "evidence", "deliverables"]:
        (ROOT / name).mkdir(parents=True, exist_ok=True)
    h = json.loads((PRIOR / "data/processed/historical_financials.json").read_text())
    original_cases = json.loads((PRIOR / "data/processed/assumptions.json").read_text())
    original_forecast = json.loads((PRIOR / "data/processed/forecast_results.json").read_text())
    market = json.loads((ROOT / "data/processed/market_inputs.json").read_text())
    for file in (PRIOR / "data/raw/reports/originals").glob("*.pdf"):
        shutil.copy2(file, ROOT / "data/raw/reports/originals" / file.name)
    for file in (PRIOR / "data/raw/source_snapshots").iterdir():
        if file.is_file(): shutil.copy2(file, ROOT / "data/raw/source_snapshots" / file.name)
    for file in ["source_manifest.json", "raw_materials_manifest.json", "selected_sec_facts.json", "primary_statement_tables.json"]:
        shutil.copy2(PRIOR / "data/raw" / file, ROOT / "data/raw" / ("prior_" + file))
    for file in ["historical_financials.json", "assumptions.json", "forecast_results.json"]:
        shutil.copy2(PRIOR / "data/processed" / file, ROOT / "data/processed" / ("prior_" + file))
    latest, notes = h["actuals"]["2026"], h["notes"]["2026"]
    pdf = fitz.open(ROOT / "data/raw/reports/originals/Microsoft_FY2026_10K.pdf")
    cover = pdf[0].get_text()
    point_shares = int(re.search(r"As of July 23, 2026, there were ([0-9,]+) shares", cover).group(1).replace(",", "")) / 1e6
    debt_text = pdf[71].get_text()
    fair_debt = float(re.search(r"estimated fair value.*?was \$([0-9.]+) billion", debt_text, re.S).group(1)) * 1000
    fraction = (dt.date(2027, 6, 30) - dt.date(2026, 10, 8)).days / 365
    historical = []
    for category, metric, label in FIELDS:
        values = []
        for year in h["historical_years"]:
            key = str(year)
            if category == "notes": amount = h["notes"][key].get(metric)
            elif category == "segments":
                index = {"pbp_revenue": 0, "ic_revenue": 1, "mpc_revenue": 2}[metric]
                segment = ["Productivity and Business Processes", "Intelligent Cloud", "More Personal Computing"][index]
                amount = h["segments"][key]["data"][segment]["revenue"]
            else: amount = h["actuals"][key][category].get(metric, 0)
            values.append(amount)
        historical.append({"category": category, "metric": metric, "label": label, "values": values})
    cases = {}
    initial_nwc = nwc(latest["balance"], notes)
    for name in ["Base", "Upside", "Downside"]:
        p = original_cases["cases"][name]
        rows = original_forecast[name]
        case = {k: p[k] for k in ["pbp_growth", "ic_growth", "mpc_growth", "tax_rate", "cash_capex_ratio", "finance_lease_add_ratio"]}
        case["ebit_margin"] = [x["operating_income"] / x["revenue"] for x in rows]
        case["da_ratio"] = [(x["depreciation"] + x["amortization"]) / x["revenue"] for x in rows]
        case["capitalized_interest_ratio"] = [x["capitalized_interest"] / x["revenue"] for x in rows]
        case["nwc_ratio"] = [(x["balance"]["receivables"] + x["balance"]["inventory"] + x["balance"]["other_current_assets"]
                                - (x["balance"]["payables"] - x["capex_payables"]) - x["balance"]["compensation"]
                                - x["balance"]["current_deferred_rev"]
                                - (x["balance"]["other_current_liab"] - x["fin_current"] - x["op_current"])) / x["revenue"] for x in rows]
        case["terminal_growth"] = {"Base": .03, "Upside": .035, "Downside": .025}[name]
        case["terminal_roic"] = {"Base": .20, "Upside": .25, "Downside": .15}[name]
        case["terminal_margin"] = {"Base": .40, "Upside": .45, "Downside": .35}[name]
        cases[name] = case
    initial = {"revenue": latest["income"]["revenue"],
               "segment_revenue": [historical[-3 + i]["values"][-1] for i in range(3)], "nwc": initial_nwc,
               "historical_ebitda": latest["income"]["operating_income"] + notes["ppe_depreciation"] + notes["intangible_amortization"]}
    controls = {"case_number": 1, "point_shares": point_shares,
                "incremental_share_proxy": latest["income"]["shares_diluted"] - latest["income"]["shares_basic"],
                "liquid_assets": latest["balance"]["cash"] + latest["balance"]["st_investments"],
                "operating_cash_reserve": 25000.0, "equity_investments": latest["balance"]["equity_investments"],
                "investment_haircut_factor": 1.0, "debt_fair_value": fair_debt,
                "finance_lease_liability": notes["finance_lease_liab"], "marginal_tax": .21,
                "terminal_tax": .20, "stub_fraction": fraction, "timing_mode": 1,
                "beta_mode": 1, "manual_beta": 1.0, "wacc_override": None}
    inputs = {"company": "Microsoft Corporation", "units": "USD millions, shares in millions",
              "valuation_date": "2026-10-08", "forecast_years": [2027, 2028, 2029, 2030, 2031],
              "market": market, "controls": controls, "initial": initial, "cases": cases,
              "historical": historical, "historical_years": h["historical_years"],
              "historical_nwc": [None] * 9 + [initial_nwc]}
    prices = list(csv.DictReader((ROOT / "data/processed/market_prices.csv").open()))
    stats = regression(prices)
    inputs["beta_results"] = {k: v for k, v in stats.items() if not isinstance(v, list)}
    (ROOT / "data/processed/model_inputs.json").write_text(json.dumps(inputs, indent=2), encoding="utf-8")
    (ROOT / "data/processed/beta_diagnostics.json").write_text(json.dumps(stats, indent=2), encoding="utf-8")
    uniforms = np.random.default_rng(20261008).random((1000, 5))
    np.savetxt(ROOT / "data/processed/monte_carlo_uniforms.csv", uniforms, delimiter=",", fmt="%.17g")
    w = wacc(inputs, stats["beta"])
    output = {"wacc": w, "bridge": bridge(inputs), "cases": {name: value(inputs, name, w["selected"]) for name in cases},
              "reverse": {name: reverse(inputs, name, w["selected"]) for name in cases},
              "monte_carlo": monte_carlo(inputs, "Base", w["selected"], uniforms)}
    (ROOT / "data/processed/reference_results.json").write_text(json.dumps(output, indent=2), encoding="utf-8")
    old_crosswalk = list(csv.DictReader((PRIOR / "docs/source_crosswalk.csv").open(encoding="utf-8-sig")))
    crosswalk = []
    for i, field in enumerate(historical):
        for j, year in enumerate(h["historical_years"]):
            candidates = [x for x in old_crosswalk if x["category"] == field["category"] and x["metric"] == field["metric"] and int(x["fiscal_year"]) == year]
            if field["category"] == "segments":
                candidates = [x for x in old_crosswalk if x["category"] == "segments" and x["metric"] == "revenue" and x["source_label"] == field["label"] and int(x["fiscal_year"]) == year]
            cell = f"Historical Inputs!{chr(69 + j)}{9 + i}"
            if candidates:
                row = dict(candidates[0]); row["excel_cells"] = cell
                row["workbook_value"] = "n.a." if field["values"][j] is None else str(field["values"][j])
                crosswalk.append(row)
            else:
                crosswalk.append({"source_id": f"DERIVED-{i}-{year}", "category": field["category"], "metric": field["metric"],
                    "fiscal_year": str(year), "report_year": "", "workbook_value": "n.a." if field["values"][j] is None else str(field["values"][j]),
                    "source_value": "", "source_label": field["label"], "status": "unavailable_disclosure",
                    "excel_cells": cell, "original_pdf": "", "highlighted_pdf": "", "pdf_page": "", "source_url": "", "unit": "USD millions"})
    crosswalk += [
        {"source_id": "DCF-SHARES", "category": "notes", "metric": "point_shares", "fiscal_year": "2026", "report_year": "2026", "source_label": "Shares outstanding as of July 23, 2026", "source_value": "7425545491", "workbook_value": str(point_shares), "unit": "million shares", "conversion": "shares divided by 1000000", "excel_cells": "Market Inputs!E11", "original_pdf": "data/raw/reports/originals/Microsoft_FY2026_10K.pdf", "highlighted_pdf": "data/raw/reports/highlighted/Microsoft_FY2026_10K_highlighted.pdf", "pdf_page": "1", "status": "reported", "source_url": "https://www.sec.gov/Archives/edgar/data/789019/000119312526323660/msft-20260630.htm"},
        {"source_id": "DCF-DEBT-FV", "category": "notes", "metric": "debt_fair_value", "fiscal_year": "2026", "report_year": "2026", "source_label": "Estimated fair value of long-term debt including current portion", "source_value": "36.5", "workbook_value": str(fair_debt), "unit": "USD millions", "conversion": "USD billions multiplied by 1000", "excel_cells": "Market Inputs!E14", "original_pdf": "data/raw/reports/originals/Microsoft_FY2026_10K.pdf", "highlighted_pdf": "data/raw/reports/highlighted/Microsoft_FY2026_10K_highlighted.pdf", "pdf_page": "72", "status": "reported_billion_to_million", "source_url": "https://www.sec.gov/Archives/edgar/data/789019/000119312526323660/msft-20260630.htm"}]
    all_keys = list(old_crosswalk[0])
    with (ROOT / "docs/source_crosswalk.csv").open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.DictWriter(stream, fieldnames=all_keys, extrasaction="ignore"); writer.writeheader(); writer.writerows(crosswalk)
    print(json.dumps({"beta": stats["beta"], "wacc": w["selected"], "nwc_initial": initial_nwc,
                       "point_shares": point_shares, "debt_fair_value": fair_debt,
                       "case_prices": {name: out["fair_price"] for name, out in output["cases"].items()},
                       "crosswalk_records": len(crosswalk)}, indent=2))


if __name__ == "__main__": main()
