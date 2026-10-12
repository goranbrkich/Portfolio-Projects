"""Reassemble the portable ZIP from GitHub download_parts with hash checks."""
import argparse
import hashlib
import json
from pathlib import Path
def main():
    parser=argparse.ArgumentParser();parser.add_argument('--parts',type=Path,default=Path(__file__).resolve().parents[1]/'download_parts');parser.add_argument('--output',type=Path,default=Path('Microsoft_DCF_ReverseDCF_Project02.zip'));a=parser.parse_args();m=json.loads((a.parts/'ARCHIVE_PARTS.json').read_text())
    for p in m['parts']:
        f=a.parts/p['name'];assert f.stat().st_size==p['bytes'] and hashlib.sha256(f.read_bytes()).hexdigest()==p['sha256'],f'Invalid archive part {f}'
    with a.output.open('wb') as out:
        for p in m['parts']:out.write((a.parts/p['name']).read_bytes())
    assert a.output.stat().st_size==m['archive_bytes'] and hashlib.sha256(a.output.read_bytes()).hexdigest()==m['archive_sha256']
    print(a.output.resolve())
if __name__=='__main__':main()
