"""Collect review evidence without modifying runs or adding a runtime dependency.

Primary --run inputs require summary.json status == "PASS". Reference inputs
are preserved transparently and never counted as accepted complete runs.
The archive contains raw logs, input/context/proof .bin files and test workbooks;
DLLs are recorded by name/hash only. No Excel or numerical solver is launched.
"""
from __future__ import annotations

import argparse
from contextlib import ExitStack
from dataclasses import dataclass
import hashlib
import json
import os
from pathlib import Path
import re
import shutil
import sys
import tempfile
from typing import Any
import zipfile


DEFAULT_ROOT = Path(__file__).resolve().parents[1]
ARCHIVE_NAME = "RPFEM_20261015_G1R19_Regression_Evidence.zip"
EMBEDDED_MANIFEST = "EVIDENCE_MANIFEST.json"
SMALL_LIMIT = 2 * 1024 * 1024
CHUNK = 1024 * 1024
SMALL_SUFFIXES = {".json", ".tsv", ".py", ".bas"}
EXCLUDED_SUFFIXES = {".dll", ".lib", ".pdb", ".exe", ".pyc", ".pyo"}
CACHE_DIRECTORIES = {
    "__pycache__", "cache", ".cache", ".pytest_cache", ".mypy_cache",
    ".ruff_cache", ".git", ".tox", ".venv",
}
LABEL = re.compile(r"[A-Za-z0-9][A-Za-z0-9_.-]{0,63}\Z")


@dataclass(frozen=True)
class Run:
    label: str
    directory: Path
    reference: bool


@dataclass(frozen=True)
class SourceFile:
    run: Run
    path: Path
    relative: str
    size: int
    sha256: str
    mtime_ns: int
    excluded: bool
    curated: bool

    @property
    def archive_path(self) -> str:
        return f"{self.run.label}/{self.relative}"


def parse_run(value: str, reference: bool) -> Run:
    label, separator, directory = value.partition("=")
    if not separator or not LABEL.fullmatch(label) or label.endswith(".") or label in {".", ".."} or not directory:
        raise ValueError("Run must be LABEL=DIR; LABEL uses 1..64 letters/digits/._- and starts with a letter/digit")
    if label.upper().split(".")[0] in {"CON", "PRN", "AUX", "NUL", *(f"COM{i}" for i in range(1, 10)), *(f"LPT{i}" for i in range(1, 10))}:
        raise ValueError(f"Windows reserved run label: {label}")
    path = Path(directory).expanduser().resolve(strict=True)
    if not path.is_dir():
        raise ValueError(f"Run is not a directory: {path}")
    return Run(label, path, reference)


def sha_file(path: Path) -> tuple[int, str]:
    digest = hashlib.sha256()
    size = 0
    with path.open("rb") as source:
        while block := source.read(CHUNK):
            digest.update(block)
            size += len(block)
    return size, digest.hexdigest()


def is_link(path: Path) -> bool:
    return path.is_symlink() or bool(getattr(path, "is_junction", lambda: False)())


def within(path: Path, directory: Path) -> bool:
    try:
        path.relative_to(directory)
        return True
    except ValueError:
        return False


def read_summary(run: Run) -> dict[str, Any]:
    path = run.directory / "summary.json"
    if is_link(path):
        raise ValueError(f"Summary cannot be a link: {path}")
    with path.open("r", encoding="utf-8-sig") as source:
        summary = json.load(source)
    if not isinstance(summary, dict):
        raise ValueError(f"Summary must be a JSON object: {path}")
    if not run.reference and summary.get("status") != "PASS":
        raise ValueError(f"Primary run {run.label} is not PASS: {summary.get('status')!r}")
    return summary


def scan_run(run: Run) -> tuple[list[SourceFile], list[str]]:
    files: list[SourceFile] = []
    skipped_directories: list[str] = []
    for directory, children, names in os.walk(run.directory, followlinks=False):
        folder = Path(directory)
        retained = []
        for name in sorted(children):
            child = folder / name
            if name.casefold() in CACHE_DIRECTORIES:
                skipped_directories.append(child.relative_to(run.directory).as_posix())
            elif is_link(child):
                raise ValueError(f"Linked directories are not collected: {child}")
            else:
                retained.append(name)
        children[:] = retained
        for name in sorted(names):
            path = folder / name
            if is_link(path) or not path.is_file() or not within(path.resolve(), run.directory):
                raise ValueError(f"Expected a regular file within the run: {path}")
            relative = path.relative_to(run.directory).as_posix()
            stat = path.stat()
            size, digest = sha_file(path)
            after = path.stat()
            if size != stat.st_size or (after.st_size, after.st_mtime_ns) != (stat.st_size, stat.st_mtime_ns):
                raise ValueError(f"Input changed while hashing: {path}")
            excluded = path.suffix.casefold() in EXCLUDED_SUFFIXES
            curated = not excluded and (relative == "summary.json" or
                (path.suffix.casefold() in SMALL_SUFFIXES and size <= SMALL_LIMIT))
            files.append(SourceFile(run, path, relative, size, digest, stat.st_mtime_ns, excluded, curated))
    return files, skipped_directories


def write_json(path: Path, value: Any) -> None:
    with path.open("w", encoding="utf-8", newline="\n") as output:
        json.dump(value, output, ensure_ascii=False, indent=2, sort_keys=True, allow_nan=False)
        output.write("\n")
        output.flush()
        os.fsync(output.fileno())


def copy_verified(source: SourceFile, destination: Path) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    with source.path.open("rb") as incoming, destination.open("xb") as outgoing:
        shutil.copyfileobj(incoming, outgoing, CHUNK)
    if sha_file(destination) != (source.size, source.sha256):
        raise ValueError(f"Copied bytes differ: {source.path}")


def verify_zip(path: Path, expected: dict[str, tuple[int, str]]) -> None:
    with zipfile.ZipFile(path, "r") as archive:
        names = archive.namelist()
        if len(names) != len(set(names)) or set(names) != set(expected):
            raise ValueError("ZIP entries do not exactly match the byte manifest")
        for name in names:
            digest = hashlib.sha256()
            size = 0
            with archive.open(name, "r") as source:
                while block := source.read(CHUNK):
                    digest.update(block)
                    size += len(block)
            if (size, digest.hexdigest()) != expected[name]:
                raise ValueError(f"ZIP byte/SHA mismatch: {name}")


def collect(runs: list[Run], root: Path, archive_path: Path) -> dict[str, Any]:
    if not runs or not any(not run.reference for run in runs):
        raise ValueError("At least one accepted primary --run is required")
    labels = [run.label.casefold() for run in runs]
    if len(labels) != len(set(labels)):
        raise ValueError("Run labels must be unique, including case")
    root = root.expanduser().resolve()
    archive_path = archive_path.expanduser().resolve()
    excel_root = root / "validation/excel"
    manifest_path = root / "validation/EXCEL_EVIDENCE_MANIFEST.json"
    validation_path = root / "delivery/EVIDENCE_ZIP_VALIDATION.json"
    destinations = [archive_path, manifest_path, validation_path, *(excel_root / run.label for run in runs)]
    for destination in destinations:
        if destination.exists():
            raise FileExistsError(f"Refusing to overwrite existing review evidence: {destination}")
        for run in runs:
            if within(destination, run.directory) or within(run.directory, destination):
                raise ValueError(f"Input/output evidence paths overlap: {run.directory} / {destination}")

    all_files: list[SourceFile] = []
    run_records = []
    for run in runs:
        summary = read_summary(run)
        files, skipped = scan_run(run)
        summary_file = next((item for item in files if item.relative == "summary.json"), None)
        if summary_file is None:
            raise ValueError(f"Run has no regular summary.json: {run.directory}")
        # Check PASS on the same immutable summary snapshot that the ZIP records.
        reread = read_summary(run)
        if reread != summary or sha_file(summary_file.path) != (summary_file.size, summary_file.sha256):
            raise ValueError(f"Summary changed during collection: {run.directory}")
        records = []
        for item in files:
            record: dict[str, Any] = dict(path=item.relative, bytes=item.size, sha256=item.sha256)
            if item.excluded:
                record["excluded_reason"] = "runtime/compiler binary or bytecode; not included in review archive"
            else:
                record["archive_path"] = item.archive_path
            if item.curated:
                record["curated_path"] = f"validation/excel/{run.label}/{item.relative}"
            records.append(record)
        run_records.append(dict(
            label=run.label, source_directory=str(run.directory),
            role="reference" if run.reference else "accepted_primary",
            scope="reference only; complete run not accepted" if run.reference else "accepted complete run; summary.status == PASS",
            summary_status_present="status" in summary, summary_status=summary.get("status"),
            accepted_complete_run=not run.reference, summary_sha256=summary_file.sha256,
            excluded_cache_directories=skipped, files=records,
            original_dll_references=[dict(filename=item.path.name, path=item.relative,
                source_path=str(item.path), bytes=item.size, sha256=item.sha256)
                for item in files if item.path.suffix.casefold() == ".dll"],
        ))
        all_files.extend(files)

    included = [item for item in all_files if not item.excluded]
    archive_names = [item.archive_path.casefold() for item in included]
    if len(archive_names) != len(set(archive_names)):
        raise ValueError("Archive paths collide, including case")
    manifest = dict(
        schema="G1R19_RELEASE_EVIDENCE_V1", release="G1R19",
        scope="review artifacts only; not required by the runtime workbook or DLL",
        source_runs_modified=False, references_are_not_accepted_complete_runs=True,
        archive_path=str(archive_path), embedded_manifest=EMBEDDED_MANIFEST,
        curated_limit_bytes=SMALL_LIMIT,
        curated_rule="summary.json plus .json/.tsv/.py/.bas <= 2 MiB; no DLL/XLSM",
        excluded_suffixes=sorted(EXCLUDED_SUFFIXES), excluded_cache_names=sorted(CACHE_DIRECTORIES),
        archived_source_files=len(included), archived_source_bytes=sum(item.size for item in included),
        runs=run_records,
    )

    # Stage owned outputs, verify them and input hashes, then publish. The final
    # validation JSON is the completion marker; a missing marker is not success.
    root.mkdir(parents=True, exist_ok=True)
    archive_path.parent.mkdir(parents=True, exist_ok=True)
    (root / "validation").mkdir(parents=True, exist_ok=True)
    (root / "delivery").mkdir(parents=True, exist_ok=True)
    with ExitStack() as stack:
        stage = Path(stack.enter_context(tempfile.TemporaryDirectory(prefix=".release-evidence-", dir=root / "validation")))
        archive_stage = Path(stack.enter_context(tempfile.TemporaryDirectory(prefix=".release-evidence-", dir=archive_path.parent)))
        staged_manifest = stage / "EXCEL_EVIDENCE_MANIFEST.json"
        write_json(staged_manifest, manifest)
        for item in all_files:
            if item.curated:
                copy_verified(item, stage / "excel" / item.run.label / item.relative)
        temporary_archive = archive_stage / ARCHIVE_NAME
        with zipfile.ZipFile(temporary_archive, "w", compression=zipfile.ZIP_DEFLATED,
                             compresslevel=6, allowZip64=True, strict_timestamps=False) as archive:
            for item in included:
                archive.write(item.path, item.archive_path)
            archive.write(staged_manifest, EMBEDDED_MANIFEST)
        expected = {item.archive_path: (item.size, item.sha256) for item in included}
        expected[EMBEDDED_MANIFEST] = sha_file(staged_manifest)
        verify_zip(temporary_archive, expected)
        # Every non-cache input file, including excluded DLL references, remains
        # byte-identical to the snapshot recorded above.
        for item in all_files:
            if sha_file(item.path) != (item.size, item.sha256):
                raise ValueError(f"Input changed before publication: {item.path}")
        for run in runs:
            expected_inventory = {item.relative for item in all_files if item.run == run}
            current_inventory = set()
            for directory, children, names in os.walk(run.directory, followlinks=False):
                folder = Path(directory)
                children[:] = [name for name in children if name.casefold() not in CACHE_DIRECTORIES]
                if any(is_link(folder / name) for name in children):
                    raise ValueError(f"Input directory links changed before publication: {folder}")
                for name in names:
                    path = folder / name
                    if is_link(path) or not path.is_file():
                        raise ValueError(f"Input file changed before publication: {path}")
                    current_inventory.add(path.relative_to(run.directory).as_posix())
            if current_inventory != expected_inventory:
                raise ValueError(f"Input file inventory changed before publication: {run.directory}")
        zip_bytes, zip_sha = sha_file(temporary_archive)
        manifest_bytes, manifest_sha = expected[EMBEDDED_MANIFEST]
        validation = dict(
            status="PASS", release="G1R19", review_only=True,
            archive=str(archive_path), archive_bytes=zip_bytes, archive_sha256=zip_sha,
            manifest="validation/EXCEL_EVIDENCE_MANIFEST.json",
            manifest_bytes=manifest_bytes, manifest_sha256=manifest_sha,
            zip_entry_count=len(expected), all_zip_file_bytes_and_sha256_verified=True,
            original_run_bytes_unchanged=True,
            accepted_primary_runs=[run.label for run in runs if not run.reference],
            reference_only_runs=[run.label for run in runs if run.reference],
            dlls_in_archive=0, xlsm_in_archive=sum(item.path.suffix.casefold() == ".xlsm" for item in included),
            completion_marker="This validation JSON is published last after all artifact checks.",
        )
        staged_validation = stage / "EVIDENCE_ZIP_VALIDATION.json"
        write_json(staged_validation, validation)
        excel_root.mkdir(parents=True, exist_ok=True)
        for run in runs:
            os.replace(stage / "excel" / run.label, excel_root / run.label)
        os.replace(temporary_archive, archive_path)
        os.replace(staged_manifest, manifest_path)
        os.replace(staged_validation, validation_path)
    return validation


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--run", action="append", default=[], metavar="LABEL=DIR", help="Accepted complete run; summary.status must be PASS")
    parser.add_argument("--reference-run", action="append", default=[], metavar="LABEL=DIR", help="Reference only; preserve status verbatim without accepting the complete run")
    parser.add_argument("--output-root", type=Path, default=DEFAULT_ROOT)
    parser.add_argument("--archive", type=Path, help=f"Default: OUTPUT_ROOT/delivery/{ARCHIVE_NAME}")
    args = parser.parse_args(argv)
    try:
        runs = [parse_run(value, False) for value in args.run]
        runs += [parse_run(value, True) for value in args.reference_run]
        output_root = args.output_root.expanduser().resolve()
        archive = args.archive or output_root / "delivery" / ARCHIVE_NAME
        result = collect(runs, output_root, archive)
    except (OSError, ValueError, zipfile.BadZipFile) as error:
        parser.exit(1, f"Evidence collection failed: {error}\n")
    print(json.dumps(result, ensure_ascii=False, sort_keys=True, allow_nan=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
