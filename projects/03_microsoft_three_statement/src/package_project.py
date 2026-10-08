"""Package the portable project with a SHA-256 manifest (standard library only)."""

import argparse
import hashlib
import json
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile


ROOT = Path(__file__).resolve().parents[1]


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    output = args.output.resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    files = sorted(
        path for path in ROOT.rglob("*")
        if path.is_file()
        and not any(part.startswith(".") or part == "__pycache__"
                    for part in path.relative_to(ROOT).parts)
        and path.suffix not in {".pyc", ".zip"}
        and path.name != "PACKAGE_MANIFEST.json"
    )
    manifest = {
        "project": "Microsoft Three Statement Model",
        "information_cutoff": "2026-10-08",
        "hash_algorithm": "SHA-256",
        "files": [
            {"path": path.relative_to(ROOT).as_posix(),
             "size_bytes": path.stat().st_size,
             "sha256": hashlib.sha256(path.read_bytes()).hexdigest()}
            for path in files
        ],
    }
    prefix = "Microsoft_Three_Statement_Project03"
    with ZipFile(output, "w", ZIP_DEFLATED, compresslevel=9) as archive:
        for path in files:
            archive.write(path, f"{prefix}/{path.relative_to(ROOT).as_posix()}")
        archive.writestr(
            f"{prefix}/PACKAGE_MANIFEST.json",
            json.dumps(manifest, indent=2) + "\n",
        )
    with ZipFile(output) as archive:
        if archive.testzip() is not None:
            raise RuntimeError("Archive CRC verification failed")
        for item in manifest["files"]:
            actual = hashlib.sha256(archive.read(f"{prefix}/{item['path']}")).hexdigest()
            if actual != item["sha256"]:
                raise RuntimeError(f"Hash mismatch: {item['path']}")
    print(json.dumps({
        "archive": str(output), "project_files": len(files),
        "size_bytes": output.stat().st_size,
        "sha256": hashlib.sha256(output.read_bytes()).hexdigest(),
        "crc_and_file_hash_checks": "passed",
    }, indent=2))


if __name__ == "__main__":
    main()
