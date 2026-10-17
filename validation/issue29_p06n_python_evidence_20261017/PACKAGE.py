"""Build once or check a self-contained diagnostic ZIP; no numerical solve."""
from pathlib import Path, PurePosixPath
import argparse
import hashlib
import json
import zipfile

ROOT = Path(__file__).resolve().parent
SIDECARS = {"PACKAGE.json", "SHA256SUMS.txt", "FINAL_ZIP_VERIFICATION.json",
            "PUBLIC_DOWNLOAD_VERIFICATION.json"}


def sha(raw):
    return hashlib.sha256(raw).hexdigest()


def files(ignore):
    result = {}
    for path in sorted(ROOT.rglob("*")):
        if not path.is_file():
            continue
        relative = path.relative_to(ROOT)
        if "__pycache__" in relative.parts or relative.parts[0] in ignore:
            continue
        if len(relative.parts) == 1 and (path.name in SIDECARS or path.suffix == ".zip"):
            continue
        result[relative.as_posix()] = path.read_bytes()
    return result


def check():
    record = json.loads((ROOT / "PACKAGE.json").read_text(encoding="utf8"))
    archive = ROOT / record["archive"]
    assert archive.parent == ROOT and archive.suffix == ".zip"
    assert archive.stat().st_size == record["bytes"] and sha(archive.read_bytes()) == record["sha256"]
    current = files(set(record["ignored_top_level_directories"]))
    assert set(current) == set(record["source_members"])
    for name, raw in current.items():
        assert sha(raw) == record["source_members"][name], name
    prefix = record["prefix"]
    with zipfile.ZipFile(archive) as z:
        assert z.testzip() is None
        assert set(z.namelist()) == {prefix + name for name in [*current, "SHA256SUMS.txt"]}
        for name, raw in current.items():
            assert z.read(prefix + name) == raw, name
        sums = (ROOT / "SHA256SUMS.txt").read_bytes()
        assert sums == z.read(prefix + "SHA256SUMS.txt") and sha(sums) == record["sha256sums_sha256"]
    return dict(status="PASS", archive=record["archive"], bytes=record["bytes"],
                sha256=record["sha256"], member_count=len(current) + 1,
                local_and_archived_bytes_identical=True, crc="PASS")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--build", metavar="ARCHIVE.zip")
    parser.add_argument("--check", action="store_true")
    parser.add_argument("--ignore-top", action="append", default=[])
    args = parser.parse_args()
    if args.check:
        assert args.build is None
        print(json.dumps(check()))
        return
    assert args.build and PurePosixPath(args.build).name == args.build and args.build.endswith(".zip")
    archive = ROOT / args.build
    assert not archive.exists() and not (ROOT / "PACKAGE.json").exists() and not (ROOT / "SHA256SUMS.txt").exists(), "Use a new package; existing evidence is immutable"
    for name in args.ignore_top:
        assert PurePosixPath(name).name == name and name not in {".", ".."}
    contents = files(set(args.ignore_top))
    prefix = ROOT.name + "/"
    sums = "".join(sha(raw) + "  " + name + "\n" for name, raw in sorted(contents.items())).encode("utf8")
    (ROOT / "SHA256SUMS.txt").write_bytes(sums)
    with zipfile.ZipFile(archive, "x", zipfile.ZIP_DEFLATED, compresslevel=6) as z:
        for name, raw in sorted(contents.items()):
            z.writestr(prefix + name, raw)
        z.writestr(prefix + "SHA256SUMS.txt", sums)
    record = dict(status="PASS", archive=archive.name, prefix=prefix,
                  bytes=archive.stat().st_size, sha256=sha(archive.read_bytes()),
                  ignored_top_level_directories=args.ignore_top,
                  source_members={name: sha(raw) for name, raw in contents.items()},
                  sha256sums_sha256=sha(sums), scope="Code/data diagnostic package; no product binary or new numerical result")
    (ROOT / "PACKAGE.json").write_text(json.dumps(record, indent=2) + "\n", encoding="utf8")
    print(json.dumps(check()))


if __name__ == "__main__":
    main()
