"""Microsoft historical OLS and HC3 analysis. Copy this entire file from Excel.

Install dependencies: python -m pip install numpy scipy matplotlib
Run: python regression_analysis.py --out regression_output
Optional refreshed data: --csv input.csv (year,revenue,cfo columns).
"""

import argparse
import csv
import json
from pathlib import Path

import numpy as np
from scipy.stats import t


# USD millions. FY2017-FY2026 source extracts, cutoff 8 October 2026.
# Revenue: Historical Inputs row 9. CFO: Historical Inputs row 80.
DATA = [
    (2017, 96571, 39507), (2018, 110360, 43884),
    (2019, 125843, 52185), (2020, 143015, 60675),
    (2021, 168088, 76740), (2022, 198270, 89035),
    (2023, 211915, 87582), (2024, 245122, 118548),
    (2025, 281724, 136162), (2026, 331839, 182935),
]
METHOD_SOURCE = (
    "https://www.statsmodels.org/stable/generated/"
    "statsmodels.regression.linear_model.OLSResults.HC3_se.html"
)


def fit_ols_hc3(x, y, years):
    """An intercept and one slope. Student t inference uses n-2 degrees of freedom."""
    x, y = np.asarray(x, dtype=float), np.asarray(y, dtype=float)
    n = len(x)
    if n < 4 or len(y) != n or not np.all(np.isfinite(x + y)):
        raise ValueError("Need at least four finite matching observations")
    design = np.column_stack([np.ones(n), x])
    if np.linalg.matrix_rank(design) != 2:
        raise ValueError("Regressor is constant")
    beta = np.linalg.lstsq(design, y, rcond=None)[0]
    fitted = design @ beta
    residual = y - fitted
    bread = np.linalg.inv(design.T @ design)
    leverage = np.einsum("ij,jk,ik->i", design, bread, design)
    if np.max(leverage) >= 1:
        raise ValueError("HC3 is undefined when leverage reaches one")
    weights = (residual / (1 - leverage)) ** 2
    covariance = bread @ (design.T @ (weights[:, None] * design)) @ bread
    df = n - 2
    rss = float(residual @ residual)
    sst = float((y - y.mean()) @ (y - y.mean()))
    if sst == 0 or rss == 0:
        raise ValueError("Inference requires variation and nonzero residuals")
    se_hc3 = np.sqrt(np.diag(covariance))
    se_ols = np.sqrt(np.diag(bread) * rss / df)
    statistics = beta / se_hc3
    pvalues = 2 * t.sf(np.abs(statistics), df)
    critical = float(t.ppf(0.975, df))
    confidence = np.column_stack([beta - critical * se_hc3,
                                  beta + critical * se_hc3])

    # Independently verify the scalar centered formulas used in Excel.
    dx = x - x.mean()
    sxx = float(dx @ dx)
    vb = float((dx ** 2) @ weights) / sxx ** 2
    va = (float(weights.sum()) / n ** 2 + x.mean() ** 2 * vb
          - 2 * x.mean() * float(dx @ weights) / (n * sxx))
    np.testing.assert_allclose([va, vb], np.diag(covariance),
                               rtol=1e-9, atol=1e-12)
    rows = []
    for i, year in enumerate(years):
        rows.append(dict(year=int(year), x=float(x[i]), y=float(y[i]),
                         fitted=float(fitted[i]), residual=float(residual[i]),
                         leverage=float(leverage[i]), hc3_weight=float(weights[i])))
    return dict(n=n, df=df, intercept=float(beta[0]), slope=float(beta[1]),
                se_ols=se_ols.tolist(), se_hc3=se_hc3.tolist(),
                t_hc3=statistics.tolist(), p_hc3=pvalues.tolist(),
                ci95_hc3=confidence.tolist(), critical_t95=critical,
                r_squared=1 - rss / sst,
                adjusted_r_squared=1 - (rss / df) / (sst / (n - 1)),
                rss=rss, sst=sst, residual_standard_error=(rss / df) ** 0.5,
                durbin_watson=float(np.diff(residual) @ np.diff(residual) / rss),
                max_leverage=float(leverage.max()),
                leverage_reference=4 / n, observations=rows)


def write_csv(path, rows):
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(rows[0]))
        writer.writeheader()
        writer.writerows(rows)


def analyze(data, out):
    data = sorted(data)
    years, revenue, cfo = np.asarray(data, dtype=float).T
    if len(set(years)) != len(years) or np.any(np.diff(years) != 1):
        raise ValueError("Need one observation per consecutive fiscal year")
    if np.any(revenue <= 0) or np.any(cfo <= 0):
        raise ValueError("Log regressions require positive revenue and CFO")
    x, y = np.log(revenue), np.log(cfo)
    levels = fit_ols_hc3(x, y, years)
    changes = fit_ols_hc3(np.diff(x), np.diff(y), years[1:])
    result = dict(
        company="Microsoft Corporation", units="USD millions",
        information_cutoff="2026-10-08", method_source=METHOD_SOURCE,
        inference="HC3 covariance with Student t(n-2); approximate small-sample inference",
        levels=levels, first_differences=changes,
        interpretation=("Historical association only. HC3 changes standard errors, not OLS "
                        "coefficients. Trending levels can produce misleading fit. "
                        "HC3 does not correct serial correlation or structural breaks. "
                        "The regressions do not replace the three-statement forecast."),
    )
    out.mkdir(parents=True, exist_ok=True)
    (out / "regression_results.json").write_text(json.dumps(result, indent=2) + "\n")
    write_csv(out / "regression_input.csv", [
        dict(year=int(yr), revenue=float(rev), cfo=float(cf))
        for yr, rev, cf in data
    ])
    write_csv(out / "levels_diagnostics.csv", levels["observations"])
    write_csv(out / "changes_diagnostics.csv", changes["observations"])
    coefficients = []
    for model, values in [("levels", levels), ("first_differences", changes)]:
        for index, name in enumerate(["intercept", "slope"]):
            coefficients.append(dict(model=model, parameter=name,
                                     estimate=values[name], se_ols=values["se_ols"][index],
                                     se_hc3=values["se_hc3"][index],
                                     t_hc3=values["t_hc3"][index],
                                     p_hc3=values["p_hc3"][index],
                                     ci95_low=values["ci95_hc3"][index][0],
                                     ci95_high=values["ci95_hc3"][index][1]))
    write_csv(out / "coefficients.csv", coefficients)

    # Optional plots. The numerical results require only NumPy and SciPy.
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
    except ImportError:
        plt = None
    if plt is not None:
        fig, ax = plt.subplots(figsize=(7, 3.2), constrained_layout=True)
        ax.plot(years, cfo / 1000, color="#17365A", marker="o", label="Actual CFO")
        ax.plot(years, np.exp([o["fitted"] for o in levels["observations"]]) / 1000,
                color="#5B9BD5", linestyle="--", label="Fitted conditional median")
        ax.set(ylabel="USD billions", xlabel="Fiscal year")
        ax.legend(frameon=False, fontsize=8)
        ax.grid(axis="y", alpha=0.2)
        fig.savefig(out / "actual_and_fitted_cfo.png", dpi=180)
        plt.close(fig)
    for name, values in [("Log levels", levels), ("Log changes", changes)]:
        print(f"{name}: n={values['n']}, slope={values['slope']:.6f}, "
              f"HC3 SE={values['se_hc3'][1]:.6f}, p={values['p_hc3'][1]:.6f}, "
              f"R2={values['r_squared']:.6f}")
    return result


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--csv", type=Path)
    parser.add_argument("--out", type=Path, default=Path("regression_output"))
    args = parser.parse_args()
    data = DATA
    if args.csv:
        with args.csv.open(newline="", encoding="utf-8-sig") as handle:
            data = [(int(r["year"]), float(r["revenue"]), float(r["cfo"]))
                    for r in csv.DictReader(handle)]
    analyze(data, args.out)


if __name__ == "__main__":
    main()
