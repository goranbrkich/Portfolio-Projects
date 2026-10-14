"""Package complete sources, deliverables and evidence; exclude internal previews and caches."""
from pathlib import Path
import json,zipfile,hashlib
ROOT=Path(__file__).resolve().parents[1]
def main():
 p=ROOT/'CHECKPOINT_STATUS.md'
 if p.exists():p.unlink()
 reports=json.loads((ROOT/'data/processed/report_manifest.json').read_text());allowed_reports={m['path'] for m in reports}
 output=ROOT.parent/'Bank_of_America_CCA_Project26.zip';paths=[]
 for p in sorted(ROOT.rglob('*')):
  if not p.is_file():continue
  rel=p.relative_to(ROOT);name=str(rel)
  if '__pycache__' in rel.parts or 'ix_cache' in rel.parts or 'pdf_checkpoints' in rel.parts:continue
  if name.startswith('qa/') and p.suffix in ['.png','.jpg','.ndjson']:continue
  if p.name=='html_highlight_checkpoint.json' or p.name.endswith('.inspect.ndjson'):continue
  if name.startswith('data/raw/reports/') and p.suffix=='.pdf' and name not in allowed_reports:continue
  paths.append(p)
 with zipfile.ZipFile(output,'w',zipfile.ZIP_DEFLATED,compresslevel=6,allowZip64=True) as z:
  for i,p in enumerate(paths):
   z.write(p,str(Path(ROOT.name)/p.relative_to(ROOT)))
   if i%100==0:print('Packaged',i,'of',len(paths),flush=True)
 with zipfile.ZipFile(output) as z:
  bad=z.testzip()
  if bad:raise ValueError('Damaged archive member: '+bad)
 info={'path':str(output),'members':len(paths),'bytes':output.stat().st_size,'sha256':hashlib.sha256(output.read_bytes()).hexdigest()};(ROOT.parent/'archive_receipt.json').write_text(json.dumps(info,indent=2));print(json.dumps(info),flush=True)
if __name__=='__main__':main()
