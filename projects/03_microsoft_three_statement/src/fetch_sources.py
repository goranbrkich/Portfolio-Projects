"""Download Microsoft primary sources into an explicit local cache.

No credentials are required. SEC requests use an identifying User-Agent and
remain below the SEC fair-access limit. Historical facts are filtered separately
to the model's fixed information cutoff.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import time
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timezone
from pathlib import Path
from urllib.request import Request, urlopen

PROJECT = Path(__file__).resolve().parents[1]
USER_AGENT = "GoranBrkichPortfolioResearch https://github.com/goranbrkich/Portfolio-Projects"
SOURCES = {
    "companyfacts.json": "https://data.sec.gov/api/xbrl/companyfacts/CIK0000789019.json",
    "submissions.json": "https://data.sec.gov/submissions/CIK0000789019.json",
    "ar26.html": "https://www.sec.gov/Archives/edgar/data/789019/000119312526323660/msft-20260630.htm",
    **{f"ar{year % 100:02}.html": f"https://www.microsoft.com/investor/reports/ar{year % 100:02}/index.html" for year in range(2016, 2026)},
}


def download(item: tuple[str, str], cache: Path, refresh: bool) -> dict:
    name, url = item
    path = cache / name
    if refresh or not path.exists():
        for attempt in range(3):
            try:
                req = Request(url, headers={"User-Agent": USER_AGENT, "Accept-Encoding": "identity"})
                with urlopen(req, timeout=35) as response:
                    body = response.read()
                if not body or (name.endswith(".json") and not body.lstrip().startswith(b"{")):
                    raise ValueError(f"Unexpected source format: {name}")
                path.write_bytes(body)
                break
            except Exception:
                if attempt == 2:
                    raise
                time.sleep(1 + attempt)
    body = path.read_bytes()
    return {"file": name, "url": url, "bytes": len(body), "sha256": hashlib.sha256(body).hexdigest(),
            "retrieved_at": datetime.now(timezone.utc).isoformat(), "source_type": "primary filing"}


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--cache", type=Path, default=PROJECT / ".cache")
    parser.add_argument("--refresh", action="store_true")
    args = parser.parse_args()
    args.cache.mkdir(parents=True, exist_ok=True)
    with ThreadPoolExecutor(max_workers=4) as executor:
        records = list(executor.map(lambda item: download(item, args.cache, args.refresh), SOURCES.items()))
    target = PROJECT / "data/raw/source_manifest.json"
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(json.dumps({"information_cutoff": "2026-10-08", "company": "Microsoft Corporation",
                                 "sources": records}, indent=2) + "\n")
    for record in records:
        print(record["file"], record["bytes"], "bytes")


if __name__ == "__main__":
    main()
