"""Reassemble the numbered GitHub archive parts and verify their hashes."""
from pathlib import Path
import json,hashlib,shutil
root=Path(__file__).resolve().parent
manifest_path=root/'archive_parts/manifest.json'
if not manifest_path.exists():raise SystemExit('Archive parts are not present. This utility is for the GitHub download; an extracted full project needs no reassembly.')
manifest=json.loads(manifest_path.read_text());output=root/manifest['archive_name']
with output.open('wb') as out:
 for item in manifest['parts']:
  p=root/'archive_parts'/item['name'];data=p.read_bytes()
  if len(data)!=item['bytes'] or hashlib.sha256(data).hexdigest()!=item['sha256']:raise SystemExit('Missing or damaged part: '+item['name'])
  out.write(data)
digest=hashlib.sha256(output.read_bytes()).hexdigest()
if digest!=manifest['sha256']:raise SystemExit('Reassembled archive hash mismatch')
print('Verified full archive:',output)
