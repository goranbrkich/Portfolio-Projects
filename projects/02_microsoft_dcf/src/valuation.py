"""Microsoft DCF, reverse DCF, OLS/HC3 and deterministic Monte Carlo.

Amounts are USD millions and shares are millions. Raw observations end on
8 October 2026. This file is also reproduced line by line in the Excel workbook.
"""
from __future__ import annotations

import argparse
import csv
import json
import math
from pathlib import Path
import xml.etree.ElementTree as ET
import zipfile

import numpy as np
from scipy import stats


def project_root() -> Path:
    for location in [Path(__file__).resolve().parent, Path.cwd()]:
        for candidate in [location, *location.parents]:
            if (candidate / "data/processed/model_inputs.json").exists():
                return candidate
    raise FileNotFoundError("Keep data/processed beside your project, then run this script from VSCode.")


def regression(prices: list[dict]) -> dict:
    """Market model with intercept, plus OLS, HC3 and Newey-West lag 5."""
    stock = np.array([float(x["msft_adjusted_close"]) for x in prices])
    market = np.array([float(x["spy_adjusted_close"]) for x in prices])
    y, x = stock[1:] / stock[:-1] - 1, market[1:] / market[:-1] - 1
    X = np.column_stack([np.ones(len(x)), x])
    inverse = np.linalg.inv(X.T @ X)
    b = inverse @ X.T @ y
    errors = y - X @ b
    leverage = np.einsum("ij,jk,ik->i", X, inverse, X)
    meat_hc3 = X.T @ ((errors / (1 - leverage))[:, None] ** 2 * X)
    cov_hc3 = inverse @ meat_hc3 @ inverse
    cov_ols = inverse * (errors @ errors / (len(x) - 2))
    scores = X * errors[:, None]
    meat_hac = scores.T @ scores
    for lag in range(1, 6):
        gamma = scores[lag:].T @ scores[:-lag]
        meat_hac += (1 - lag / 6) * (gamma + gamma.T)
    cov_hac = inverse @ meat_hac @ inverse * len(x) / (len(x) - 2)
    critical = float(stats.t.ppf(0.975, len(x) - 2))
    sxx = float(np.sum((x - x.mean()) ** 2))
    sse, sst = float(errors @ errors), float(np.sum((y - y.mean()) ** 2))
    se = np.sqrt(np.diag(cov_hc3))
    return {"n": len(x), "alpha": float(b[0]), "beta": float(b[1]),
            "beta_blume": float(2 / 3 * b[1] + 1 / 3), "r_squared": 1 - sse / sst,
            "ols_se_alpha": float(np.sqrt(cov_ols[0, 0])), "ols_se_beta": float(np.sqrt(cov_ols[1, 1])),
            "hc3_se_alpha": float(se[0]), "hc3_se_beta": float(se[1]),
            "hac5_se_beta": float(np.sqrt(cov_hac[1, 1])), "critical_t": critical,
            "hc3_beta_lower": float(b[1] - critical * se[1]),
            "hc3_beta_upper": float(b[1] + critical * se[1]),
            "hc3_beta_p_value": float(2 * stats.t.sf(abs(b[1] / se[1]), len(x) - 2)),
            "sxx": sxx, "sse": sse, "sst": sst,
            "mean_x": float(x.mean()), "mean_y": float(y.mean()),
            "x": x.tolist(), "y": y.tolist(), "residuals": errors.tolist(),
            "leverage": leverage.tolist()}


def wacc(inputs: dict, beta: float) -> dict:
    control, market = inputs["controls"], inputs["market"]
    selected_beta = {1: beta, 2: 2 / 3 * beta + 1 / 3, 3: control["manual_beta"]}[int(control["beta_mode"])]
    equity = market["price"] * control["point_shares"]
    debt = control["debt_fair_value"] + control["finance_lease_liability"]
    equity_weight = equity / (equity + debt)
    ke = market["riskfree"] + selected_beta * market["erp"]
    kd = market["riskfree"] + market["debt_spread"]
    calculated = equity_weight * ke + (1 - equity_weight) * kd * (1 - control["marginal_tax"])
    chosen = calculated if control["wacc_override"] is None else control["wacc_override"]
    return {"beta": selected_beta, "ke": ke, "kd": kd, "equity_weight": equity_weight,
            "debt_weight": 1 - equity_weight, "calculated": calculated, "selected": chosen,
            "market_equity": equity, "market_debt": debt}


def bridge(inputs: dict) -> dict:
    c = inputs["controls"]
    excess_cash = max(c["liquid_assets"] - c["operating_cash_reserve"], 0)
    investments = c["equity_investments"] * c["investment_haircut_factor"]
    adjustment = excess_cash + investments - c["debt_fair_value"] - c["finance_lease_liability"]
    shares = c["point_shares"] + c["incremental_share_proxy"]
    target = inputs["market"]["price"] * shares - adjustment
    return {"excess_cash": excess_cash, "investments": investments,
            "adjustment": adjustment, "shares": shares, "market_enterprise_value": target}


def forecast(inputs: dict, case: str, growth=None, margin_shift=0.0,
             capex_scale=1.0, nwc_shift=0.0, tax_shift=0.0, lease_scale=1.0) -> dict:
    p = inputs["cases"][case]
    segments = np.asarray(inputs["initial"]["segment_revenue"], dtype=float).copy()
    revenue, segment_rows = [], []
    for i in range(5):
        segments = segments * (1 + np.array([p["pbp_growth"][i], p["ic_growth"][i], p["mpc_growth"][i]]))
        segment_rows.append(segments.tolist())
        revenue.append(float(segments.sum()))
    if growth is not None:
        revenue = [inputs["initial"]["revenue"] * (1 + growth) ** (i + 1) for i in range(5)]
    revenue = np.asarray(revenue)
    margin = np.asarray(p["ebit_margin"]) + margin_shift
    tax = np.asarray(p["tax_rate"]) + tax_shift
    ebit = revenue * margin
    nopat = ebit * (1 - tax)
    da = revenue * np.asarray(p["da_ratio"])
    cash_capex = revenue * np.asarray(p["cash_capex_ratio"]) * capex_scale
    capitalized_interest = revenue * np.asarray(p["capitalized_interest_ratio"])
    finance_additions = revenue * np.asarray(p["finance_lease_add_ratio"]) * lease_scale
    nwc = revenue * (np.asarray(p["nwc_ratio"]) + nwc_shift)
    change_nwc = np.diff(np.r_[inputs["initial"]["nwc"], nwc])
    economic_capex = cash_capex - capitalized_interest + finance_additions
    fcff = nopat + da - economic_capex - change_nwc
    return {"revenue": revenue.tolist(), "segments": segment_rows, "ebit": ebit.tolist(),
            "nopat": nopat.tolist(), "da": da.tolist(), "cash_capex": cash_capex.tolist(),
            "capitalized_interest": capitalized_interest.tolist(), "finance_additions": finance_additions.tolist(),
            "economic_capex": economic_capex.tolist(), "nwc": nwc.tolist(),
            "change_nwc": change_nwc.tolist(), "fcff": fcff.tolist()}


def value(inputs: dict, case: str, discount: float, **shock) -> dict:
    growth = shock.pop("terminal_growth", inputs["cases"][case]["terminal_growth"])
    roic = shock.pop("terminal_roic", inputs["cases"][case]["terminal_roic"])
    margin_shift = shock.get("margin_shift", 0.0)
    tax_shift = shock.get("tax_shift", 0.0)
    stable_margin = shock.pop("stable_margin", inputs["cases"][case]["terminal_margin"] + margin_shift)
    control = inputs["controls"]
    if not (discount > growth and roic > growth):
        raise ValueError("Require WACC > terminal growth and terminal ROIC > growth")
    f = forecast(inputs, case, **shock)
    fraction = control["stub_fraction"]
    weights = np.array([fraction, 1, 1, 1, 1])
    end = fraction + np.arange(5)
    mid = np.r_[fraction / 2, fraction + np.arange(4) + 0.5]
    timing = mid if int(control["timing_mode"]) == 1 else end
    df = (1 + discount) ** timing
    pv = np.asarray(f["fcff"]) * weights / df
    stable_revenue = f["revenue"][-1] * (1 + growth)
    stable_nopat = stable_revenue * stable_margin * (1 - control["terminal_tax"] - tax_shift)
    reinvestment = stable_nopat * growth / roic
    stable_fcff = stable_nopat - reinvestment
    terminal_value = stable_fcff / (discount - growth)
    terminal_df = (1 + discount) ** end[-1]
    terminal_pv = terminal_value / terminal_df
    enterprise = float(pv.sum() + terminal_pv)
    b = bridge(inputs)
    equity = enterprise + b["adjustment"]
    return {"case": case, "wacc": discount, "terminal_growth": growth, "terminal_roic": roic,
            "forecast": f, "timing": timing.tolist(), "cashflow_weights": weights.tolist(),
            "discount_factors": df.tolist(), "pv_cashflows": pv.tolist(),
            "pv_explicit": float(pv.sum()), "terminal_revenue": stable_revenue,
            "terminal_nopat": stable_nopat, "terminal_reinvestment": reinvestment,
            "terminal_fcff": stable_fcff, "terminal_value": terminal_value,
            "terminal_df": terminal_df, "pv_terminal": terminal_pv,
            "terminal_share": terminal_pv / enterprise, "enterprise_value": enterprise,
            "equity_value": equity, "fair_price": equity / b["shares"],
            "upside": equity / b["shares"] / inputs["market"]["price"] - 1}


def bisect(function, low: float, high: float, iterations=48):
    fl, fh = function(low), function(high)
    if fl * fh > 0:
        return None
    for _ in range(iterations):
        midpoint = (low + high) / 2
        fm = function(midpoint)
        if fl * fm <= 0:
            high, fh = midpoint, fm
        else:
            low, fl = midpoint, fm
    return (low + high) / 2


def reverse(inputs: dict, case: str, discount: float) -> dict:
    baseline = value(inputs, case, discount)
    target = bridge(inputs)["market_enterprise_value"]
    gap = target - baseline["enterprise_value"]
    p, c, f = inputs["cases"][case], inputs["controls"], baseline["forecast"]
    rev = np.asarray(f["revenue"])
    weighted_df = np.asarray(baseline["cashflow_weights"]) / np.asarray(baseline["discount_factors"])
    terminal_coefficient = baseline["terminal_revenue"] * (1 - c["terminal_tax"]) * (1 - p["terminal_growth"] / p["terminal_roic"]) / ((discount - p["terminal_growth"]) * baseline["terminal_df"])
    margin_slope = float(np.sum(rev * (1 - np.asarray(p["tax_rate"])) * weighted_df) + terminal_coefficient)
    capex_slope = -float(np.sum(np.asarray(f["cash_capex"]) * weighted_df))
    lease_slope = -float(np.sum(np.asarray(f["finance_additions"]) * weighted_df))
    nwc_slope = -float(np.sum(np.diff(np.r_[0.0, rev]) * weighted_df))
    tax_slope = -float(np.sum(np.asarray(f["ebit"]) * weighted_df) + baseline["terminal_revenue"] * p["terminal_margin"] * (1 - p["terminal_growth"] / p["terminal_roic"]) / ((discount - p["terminal_growth"]) * baseline["terminal_df"]))
    required_terminal_pv = target - baseline["pv_explicit"]
    required_fcff = required_terminal_pv * baseline["terminal_df"] * (discount - p["terminal_growth"])
    nopat_pv = baseline["terminal_nopat"] / ((discount - p["terminal_growth"]) * baseline["terminal_df"])
    roic_denominator = 1 - required_terminal_pv / nopat_pv
    required_roic = p["terminal_growth"] / roic_denominator if roic_denominator > 0 else None
    growth = bisect(lambda x: value(inputs, case, discount, growth=x)["enterprise_value"] - target, -0.15, 0.75)
    implied_wacc = bisect(lambda x: value(inputs, case, x)["enterprise_value"] - target, p["terminal_growth"] + 0.001, 0.30)
    implied_g = bisect(lambda x: value(inputs, case, discount, terminal_growth=x)["enterprise_value"] - target, 0.0, min(discount, p["terminal_roic"]) - 0.001)
    return {"market_enterprise_value": target, "valuation_gap": gap,
            "growth": growth, "margin_shift": gap / margin_slope,
            "terminal_margin_only": p["terminal_margin"] + gap / terminal_coefficient,
            "forecast_margin_only_shift": gap / (margin_slope - terminal_coefficient),
            "capex_scale": 1 + gap / capex_slope, "lease_scale": 1 + gap / lease_slope,
            "nwc_shift": gap / nwc_slope, "dso_change_days": 365 * gap / nwc_slope,
            "tax_shift": gap / tax_slope, "wacc": implied_wacc,
            "terminal_growth": implied_g, "terminal_roic": required_roic,
            "roic_limit_price": (baseline["pv_explicit"] + nopat_pv + bridge(inputs)["adjustment"]) / bridge(inputs)["shares"],
            "required_terminal_fcff": required_fcff,
            "required_terminal_fcff_margin": required_fcff / baseline["terminal_revenue"],
            "required_exit_multiple": required_terminal_pv * baseline["terminal_df"] / (f["ebit"][-1] + f["da"][-1]),
            "required_shares": baseline["equity_value"] / inputs["market"]["price"],
            "derivatives": {"margin": margin_slope, "capex": capex_slope, "lease": lease_slope,
                            "nwc": nwc_slope, "tax": tax_slope, "terminal_margin": terminal_coefficient}}


def triangular(u: np.ndarray, low: float, mode: float, high: float) -> np.ndarray:
    if not low <= mode <= high or not low < high:
        raise ValueError("Triangular bounds must satisfy low <= mode <= high and low < high")
    split = (mode - low) / (high - low)
    return np.where(u < split, low + np.sqrt(u * (high - low) * (mode - low)),
                    high - np.sqrt((1 - u) * (high - low) * (high - mode)))


def monte_carlo(inputs: dict, case: str, discount: float, uniforms: np.ndarray) -> dict:
    baseline = value(inputs, case, discount)
    base_cagr = (baseline["forecast"]["revenue"][-1] / inputs["initial"]["revenue"]) ** 0.2 - 1
    p = inputs["cases"][case]
    bounds = [[base_cagr - 0.05, base_cagr, base_cagr + 0.05], [-0.05, 0, 0.05],
              [0.85, 1, 1.20], [discount - 0.015, discount, discount + 0.015],
              [p["terminal_growth"] - 0.01, p["terminal_growth"], p["terminal_growth"] + 0.01]]
    draws = np.column_stack([triangular(uniforms[:, i], *bounds[i]) for i in range(5)])
    prices = np.array([value(inputs, case, float(row[3]), growth=float(row[0]),
                             margin_shift=float(row[1]), capex_scale=float(row[2]),
                             terminal_growth=float(row[4]))["fair_price"] for row in draws])
    return {"seed": 20261008, "draws": draws.tolist(), "prices": prices.tolist(), "bounds": bounds,
            "percentiles": {str(q): float(np.quantile(prices, q)) for q in [0.05, 0.25, 0.5, 0.75, 0.95]},
            "mean": float(prices.mean()), "share_above_market": float(np.mean(prices > inputs["market"]["price"]))}


def xlsx_inputs(path: Path, inputs: dict, mapping: dict) -> dict:
    """Read editable source cells, never rely on formula caches to recalculate."""
    result = json.loads(json.dumps(inputs))
    result.setdefault("historical_inputs", {})
    ns = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main", "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
    with zipfile.ZipFile(path) as archive:
        book = ET.fromstring(archive.read("xl/workbook.xml"))
        relations = {x.attrib["Id"]: x.attrib["Target"] for x in ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))}
        names = {}
        for sheet in book.find("s:sheets", ns):
            target = relations[sheet.attrib["{" + ns["r"] + "}id"]]
            names[sheet.attrib["name"]] = target.lstrip("/") if target.startswith("/") else "xl/" + target
        shared = []
        if "xl/sharedStrings.xml" in archive.namelist():
            shared = ["".join(x.itertext()) for x in ET.fromstring(archive.read("xl/sharedStrings.xml"))]
        tables = {}
        for name in {x["sheet"] for x in mapping["cells"]}:
            tables[name] = {cell.attrib["r"]: cell for cell in ET.fromstring(archive.read(names[name])).findall(".//s:c", ns)}
        for item in mapping["cells"]:
            cell = tables[item["sheet"]].get(item["cell"])
            value = None if cell is None else cell.findtext("s:v", None, ns)
            if cell is not None and cell.find("s:f", ns) is not None:
                raise ValueError(f"Editable cell {item['sheet']}!{item['cell']} contains a formula. Supply a numeric driver.")
            if value is not None:
                value = shared[int(value)] if cell.attrib.get("t") == "s" else float(value)
            if cell is not None and cell.attrib.get("t") == "inlineStr":
                value = "".join(cell.find("s:is", ns).itertext())
            target = result
            keys = item["key"]
            for key in keys[:-1]:
                target = target[key] if not isinstance(target, list) else target[int(key)]
            if isinstance(target, list):
                target[int(keys[-1])] = value
            else:
                target[keys[-1]] = value
    h = result["historical_inputs"]
    if h:
        result["initial"]["revenue"] = h["revenue"]
        result["initial"]["segment_revenue"] = [h[x] for x in ["pbp_revenue", "ic_revenue", "mpc_revenue"]]
        result["initial"]["nwc"] = (h["receivables"] + h["inventory"] + h["other_current_assets"]
            - h["payables"] + h["capex_payables"] - h["compensation"] - h["current_deferred_rev"]
            - h["other_current_liab"] + h["op_lease_current"] + h["finance_lease_current"])
        result["controls"]["liquid_assets"] = h["cash"] + h["st_investments"]
        result["controls"]["equity_investments"] = h["equity_investments"]
        result["controls"]["finance_lease_liability"] = h["finance_lease_liab"]
    return result


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--case", choices=["Base", "Upside", "Downside"], default=None)
    parser.add_argument("--excel", type=Path)
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    root = project_root()
    inputs = json.loads((root / "data/processed/model_inputs.json").read_text())
    if args.excel:
        inputs = xlsx_inputs(args.excel, inputs, json.loads((root / "data/processed/excel_input_map.json").read_text()))
    case = args.case or ["Base", "Upside", "Downside"][int(inputs["controls"]["case_number"]) - 1]
    prices = list(csv.DictReader((root / "data/processed/market_prices.csv").open()))
    beta = regression(prices)
    rates = wacc(inputs, beta["beta"])
    discount = rates["selected"]
    uniforms = np.array(list(csv.reader((root / "data/processed/monte_carlo_uniforms.csv").open())), dtype=float)
    result = {"valuation_date": inputs["market"]["valuation_date"], "selected_case": case,
              "wacc": rates, "bridge": bridge(inputs),
              "beta": {k: v for k, v in beta.items() if not isinstance(v, list)},
              "selected": value(inputs, case, discount), "reverse": reverse(inputs, case, discount),
              "cases": {name: value(inputs, name, discount) for name in inputs["cases"]},
              "monte_carlo": monte_carlo(inputs, case, discount, uniforms)}
    output = args.output or root / "data/processed/python_results.json"
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(result, indent=2), encoding="utf-8")
    print(json.dumps({"case": case, "wacc": discount, "price": result["selected"]["fair_price"],
                      "market_price": inputs["market"]["price"], "beta": beta["beta"],
                      "hc3_se_beta": beta["hc3_se_beta"], "reverse": result["reverse"]}, indent=2))


if __name__ == "__main__":
    main()
