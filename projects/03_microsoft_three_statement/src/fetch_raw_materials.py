"""Download the exact filing versions used by the supplied Microsoft model.

Native PDFs are obtained from Microsoft's linked investor-relations provider.
Raw HTML and SEC API snapshots supplement the existing frozen extracts. The
original model manifest is preserved; current downloads have their own hashes.
"""
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timezone
import hashlib
import json
from pathlib import Path
from urllib.request import Request, urlopen

import fitz

PROJECT = Path(__file__).resolve().parents[1]
HEADERS = {'User-Agent': 'GoranBrkichPortfolioResearch https://github.com/goranbrkich/Portfolio-Projects', 'Accept-Encoding': 'identity'}


def download(item):
    path = PROJECT / item['path']
    path.parent.mkdir(parents=True, exist_ok=True)
    if not path.exists():
        for attempt in range(2):
            try:
                with urlopen(Request(item['url'], headers=HEADERS), timeout=30) as response:
                    body = response.read()
                break
            except Exception:
                if attempt:
                    raise
        if item['kind'] == 'pdf' and not body.startswith(b'%PDF-'):
            raise ValueError(f'Not a PDF: {item["url"]}')
        if item['kind'] == 'json':
            json.loads(body)
        path.write_bytes(body)
    body = path.read_bytes()
    record = {**item, 'size_bytes': len(body), 'sha256': hashlib.sha256(body).hexdigest(),
              'retrieved_at': datetime.now(timezone.utc).isoformat()}
    if item['kind'] == 'pdf':
        with fitz.open(path) as doc:
            record['pdf_pages'] = len(doc)
            record['searchable_text'] = any(page.get_text().strip() for page in doc)
        if not record['searchable_text']:
            raise ValueError(f'No searchable text: {path}')
    if item.get('previous_sha256'):
        record['same_bytes_as_initial_model_source'] = record['sha256'] == item['previous_sha256']
    return record


def main():
    discovery = json.loads((PROJECT / 'evidence/official_pdf_discovery.json').read_text())
    jobs = []
    for year in discovery:
        filing = year['filings'][0]
        pdf = next(x for x in filing['links'] if x['text'].endswith('.pdf'))
        accn = pdf['text'].removesuffix('.pdf')
        jobs.append({'kind': 'pdf', 'report_year': year['report_year'], 'accession': accn,
                     'listing_url': year['listing_url'], 'url': pdf['url'],
                     'path': f'data/raw/reports/originals/Microsoft_FY{year["report_year"]}_10K.pdf'})
    original = json.loads((PROJECT / 'data/raw/source_manifest.json').read_text())
    for source in original['sources']:
        kind = 'json' if source['file'].endswith('.json') else 'html'
        jobs.append({'kind': kind, 'url': source['url'],
                     'path': 'data/raw/source_snapshots/' + source['file'],
                     'previous_sha256': source['sha256']})
    with ThreadPoolExecutor(max_workers=3) as pool:
        records = list(pool.map(download, jobs))
    manifest = {'company': 'Microsoft Corporation', 'model_information_cutoff': '2026-10-08',
                'download_date': datetime.now(timezone.utc).date().isoformat(),
                'original_model_manifest_preserved': True,
                'snapshot_note': 'New raw downloads are dated separately. The supplied selected_sec_facts.json is the frozen selection used by the workbook.',
                'sources': records}
    (PROJECT / 'data/raw/raw_materials_manifest.json').write_text(json.dumps(manifest, indent=2) + '\n')
    print(json.dumps({'pdf_reports': len(discovery), 'raw_snapshots': len(original['sources']),
                      'download_bytes': sum(r['size_bytes'] for r in records),
                      'changed_source_snapshots': [r['path'] for r in records if r.get('same_bytes_as_initial_model_source') is False],
                      'pdf_details': [(r['report_year'], r['size_bytes'], r['pdf_pages']) for r in records if r['kind'] == 'pdf']}, indent=2))


if __name__ == '__main__':
    main()
