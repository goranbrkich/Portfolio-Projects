"""Freeze market inputs and original provider responses for the 8 Oct 2026 case."""
from __future__ import annotations

import concurrent.futures
import csv
import datetime as dt
import hashlib
import html
import io
import json
import mimetypes
import re
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
RAW = ROOT / "data" / "raw" / "market"
PROCESSED = ROOT / "data" / "processed"
CUTOFF = dt.date(2026, 10, 8)
START = dt.date(2021, 10, 7)
P1 = int(dt.datetime.combine(START, dt.time(), dt.timezone.utc).timestamp())
P2 = int(dt.datetime(2026, 10, 9, tzinfo=dt.timezone.utc).timestamp())
URLS = {
    "msft_prices.json": f"https://query1.finance.yahoo.com/v8/finance/chart/MSFT?period1={P1}&period2={P2}&interval=1d&events=div%2Csplits&includeAdjustedClose=true",
    "spy_prices.json": f"https://query1.finance.yahoo.com/v8/finance/chart/SPY?period1={P1}&period2={P2}&interval=1d&events=div%2Csplits&includeAdjustedClose=true",
    "treasury_oct2026.csv": "https://home.treasury.gov/resource-center/data-chart-center/interest-rates/daily-treasury-rates.csv/all/202610?_format=csv&field_tdr_date_value_month=202610&page=&type=daily_treasury_yield_curve",
    "treasury_oct2026.html": "https://home.treasury.gov/resource-center/data-chart-center/interest-rates/TextView?type=daily_treasury_yield_curve&field_tdr_date_value=202610",
    "damodaran_home.html": "https://pages.stern.nyu.edu/~adamodar/New_Home_Page/home.htm",
    "damodaran_ratings_jan2026.html": "https://pages.stern.nyu.edu/~adamodar/New_Home_Page/datafile/ratings.html",
    "ERPbymonth.xlsx": "https://pages.stern.nyu.edu/~adamodar/pc/implprem/ERPbymonth.xlsx",
    "ERPOct26.xlsx": "https://pages.stern.nyu.edu/~adamodar/pc/implprem/ERPOct26.xlsx",
    "microsoft_stock_lookup.html": "https://microsoft.gcs-web.com/",
    "hc3_method.html": "https://www.statsmodels.org/stable/generated/statsmodels.regression.linear_model.OLSResults.HC3_se.html",
}


def fetch(item):
    name, url = item
    existing = RAW / name
    if existing.exists():
        body = existing.read_bytes()
        return {"name": name, "url": url, "path": f"data/raw/market/{name}",
                "bytes": len(body), "mime_type": mimetypes.guess_type(name)[0],
                "sha256": hashlib.sha256(body).hexdigest(),
                "retrieved_at": dt.datetime.fromtimestamp(existing.stat().st_mtime, dt.timezone.utc).isoformat()}
    request = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(request, timeout=35) as response:
        body = response.read()
        mime = response.headers.get("Content-Type", "")
    (RAW / name).write_bytes(body)
    return {"name": name, "url": url, "path": f"data/raw/market/{name}",
            "bytes": len(body), "mime_type": mime,
            "sha256": hashlib.sha256(body).hexdigest(),
            "retrieved_at": dt.datetime.now(dt.timezone.utc).isoformat()}


def quotes(name):
    response = json.loads((RAW / name).read_bytes())
    if response["chart"]["error"]:
        raise ValueError(response["chart"]["error"])
    series = response["chart"]["result"][0]
    quote = series["indicators"]["quote"][0]
    adjusted = series["indicators"]["adjclose"][0]["adjclose"]
    rows = []
    for index, timestamp in enumerate(series["timestamp"]):
        date = dt.datetime.fromtimestamp(timestamp, dt.timezone.utc).date()
        if date > CUTOFF:
            continue
        if quote["close"][index] is None or adjusted[index] is None:
            continue
        rows.append({"date": date.isoformat(), "close": quote["close"][index],
                     "adjusted_close": adjusted[index], "volume": quote["volume"][index]})
    assert rows[-1]["date"] == CUTOFF.isoformat(), "Missing cutoff quote"
    return rows


def main():
    RAW.mkdir(parents=True, exist_ok=True)
    PROCESSED.mkdir(parents=True, exist_ok=True)
    with concurrent.futures.ThreadPoolExecutor(max_workers=5) as pool:
        manifest = list(pool.map(fetch, URLS.items()))
    (RAW / "market_source_manifest.json").write_text(json.dumps(manifest, indent=2), encoding="utf-8")
    msft, spy = quotes("msft_prices.json"), quotes("spy_prices.json")
    market = {row["date"]: row for row in spy}
    rows = []
    for stock in msft:
        benchmark = market.get(stock["date"])
        if benchmark:
            rows.append({"date": stock["date"], "msft_close": stock["close"],
                         "msft_adjusted_close": stock["adjusted_close"],
                         "spy_close": benchmark["close"],
                         "spy_adjusted_close": benchmark["adjusted_close"]})
    with (PROCESSED / "market_prices.csv").open("w", newline="", encoding="utf-8") as stream:
        writer = csv.DictWriter(stream, fieldnames=list(rows[0]))
        writer.writeheader()
        writer.writerows(rows)
    treasury = list(csv.DictReader(io.StringIO((RAW / "treasury_oct2026.csv").read_text(encoding="utf-8-sig"))))
    chosen = next(row for row in treasury if row.get("Date") == "10/08/2026")
    home = (RAW / "damodaran_home.html").read_text(errors="replace")
    flat = html.unescape(re.sub("<[^>]+>", " ", home))
    match = re.search(r"Implied ERP on\s+October\s+1,\s+2026\s*=\s*([0-9.\s]+)%", flat, re.I)
    assert match, "The as-of October ERP was not present in the author source"
    inputs = {"valuation_date": CUTOFF.isoformat(), "price": round(rows[-1]["msft_close"], 2),
              "price_basis": "Unadjusted MSFT regular-session close in USD, 8 Oct 2026",
              "beta_estimation": "Five-year market model with intercept, total-return adjusted MSFT and SPY daily prices",
              "riskfree": float(chosen["10 Yr"]) / 100,
              "three_month_treasury": float(chosen["3 Mo"]) / 100,
              "erp": float(re.sub(r"\s+", "", match.group(1))) / 100,
              "erp_observation_date": "2026-10-01", "debt_spread": 0.004,
              "spread_observation_date": "2026-01-01", "treasury_cutoff_row": chosen,
              "observations": len(rows), "sources": manifest,
              "erp_period_validation": "Use the explicitly dated October 1, 2026 author homepage. ERPbymonth.xlsx has a conflicting last Excel date serial (44834, October 1, 2022) and is not used to select the valuation period."}
    (PROCESSED / "market_inputs.json").write_text(json.dumps(inputs, indent=2), encoding="utf-8")
    (RAW / "market_source_manifest.json").write_text(json.dumps(manifest, indent=2), encoding="utf-8")
    print(json.dumps({k: v for k, v in inputs.items() if k not in ["sources", "treasury_cutoff_row"]}))


if __name__ == "__main__":
    main()
