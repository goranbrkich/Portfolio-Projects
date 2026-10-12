"""Join the verified GitHub archive parts into the complete project ZIP."""
import argparse
import hashlib
import json
from pathlib import Path


def main():
    root = Path(__file__).resolve().parents[1]
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--parts', type=Path, default=root / 'archive_parts')
    parser.add_argument('--output', type=Path, default=root / 'Microsoft_Three_Statement_Project03.zip')
    args = parser.parse_args()
    manifest = json.loads((args.parts / 'ARCHIVE_PARTS.json').read_text())
    if args.output.exists():
        if hashlib.sha256(args.output.read_bytes()).hexdigest() == manifest['archive_sha256']:
            print('Complete project ZIP already matches the archive manifest.')
            return
        raise FileExistsError(f'Choose another --output path: {args.output}')
    # Validate every input before creating the output.
    for part in manifest['parts']:
        data = (args.parts / part['file']).read_bytes()
        if len(data) != part['size_bytes'] or hashlib.sha256(data).hexdigest() != part['sha256']:
            raise ValueError(f'Archive part failed verification: {part["file"]}')
    args.output.parent.mkdir(parents=True, exist_ok=True)
    with args.output.open('wb') as output:
        for part in manifest['parts']:
            output.write((args.parts / part['file']).read_bytes())
    if args.output.stat().st_size != manifest['archive_size_bytes'] or hashlib.sha256(args.output.read_bytes()).hexdigest() != manifest['archive_sha256']:
        args.output.unlink()
        raise ValueError('Reassembled archive failed verification')
    print(f'Created and verified {args.output}')


if __name__ == '__main__':
    main()
