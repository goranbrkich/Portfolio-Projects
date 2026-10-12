"""Build a portable ZIP and SHA-256 inventory from the final project."""
from pathlib import Path
import hashlib
import json
import zipfile
import argparse

ROOT=Path(__file__).resolve().parents[1]
def files():
    return [p for p in sorted(ROOT.rglob('*')) if p.is_file() and '__pycache__' not in p.parts
            and not any(x in p.parts for x in ['workbook_previews','report_previews','pdf_previews','download_parts'])
            and p.name not in ['PACKAGE_MANIFEST.json','WORK_CHECKPOINT.json','publication_state.json']]
def main():
    parser=argparse.ArgumentParser();parser.add_argument('--output',type=Path,required=True);a=parser.parse_args()
    paths=files();manifest={'project':'02 Microsoft DCF and Reverse DCF','information_cutoff':'2026-10-08',
                          'files':[{'path':str(p.relative_to(ROOT)),'bytes':p.stat().st_size,'sha256':hashlib.sha256(p.read_bytes()).hexdigest()} for p in paths]}
    m=ROOT/'PACKAGE_MANIFEST.json';m.write_text(json.dumps(manifest,indent=2)+'\n');paths.append(m)
    a.output.parent.mkdir(parents=True,exist_ok=True)
    with zipfile.ZipFile(a.output,'w',compression=zipfile.ZIP_DEFLATED,compresslevel=6) as z:
        for p in paths:z.write(p,'Microsoft_DCF_ReverseDCF_Project02/'+str(p.relative_to(ROOT)))
    with zipfile.ZipFile(a.output) as z:assert z.testzip() is None
    print(json.dumps({'file':str(a.output),'bytes':a.output.stat().st_size,'sha256':hashlib.sha256(a.output.read_bytes()).hexdigest(),'files':len(paths)}))
if __name__=='__main__':main()
