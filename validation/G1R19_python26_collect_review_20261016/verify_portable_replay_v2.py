"""Independently extract, supervise and audit a frozen saved-pair archive."""
from pathlib import Path, PurePosixPath
from datetime import datetime, timezone
import argparse
import hashlib
import json
import os
import stat
import subprocess
import sys
import time
import zipfile
import psutil

EXPECTED_SHA = '9121db1fa2f60ebdddf35d8c6123d4562cf631a3808c61cb02d2a36702e2960a'
EXPECTED_SIZE = 98287686
REPO = Path(__file__).resolve().parents[2]


def digest(path):
    result = hashlib.sha256()
    with Path(path).open('rb') as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b''):
            result.update(block)
    return result.hexdigest()


def write(path, value):
    temp = path.with_name(path.name + '.tmp')
    temp.write_text(json.dumps(value, indent=2, ensure_ascii=False, allow_nan=False) + '\n', encoding='utf8')
    temp.replace(path)


def process_record(process):
    return dict(pid=process.pid, create_time=process.create_time(),
                cmdline=process.cmdline(), exe=process.exe())


def unchanged_process(record):
    try:
        process = psutil.Process(record['pid'])
        return (process.create_time() == record['create_time'] and
                process.cmdline() == record['cmdline'])
    except (psutil.NoSuchProcess, psutil.ZombieProcess):
        return False


def inventory(root):
    return {str(p.relative_to(root)).replace('\\', '/'): digest(p)
            for p in sorted(root.rglob('*')) if p.is_file()}


def safe_extract(archive, target):
    if target.exists():
        raise ValueError('Independent extraction folder must not exist')
    names, original_hashes = set(), {}
    with zipfile.ZipFile(archive) as package:
        infos = package.infolist()
        for item in infos:
            name = item.filename
            normalized = PurePosixPath(name)
            if ('\\' in name or ':' in name or normalized.is_absolute() or '..' in normalized.parts or
                not name or item.is_dir() or name.casefold() in names or
                stat.S_ISLNK(item.external_attr >> 16)):
                raise ValueError('Unsafe or duplicate ZIP member: ' + name)
            names.add(name.casefold())
            resolved = (target / name).resolve()
            if not resolved.is_relative_to(target.resolve()):
                raise ValueError('ZIP member escaped extraction folder')
        if len(infos) != 1107:
            raise ValueError('Expected1105 evidence files plus manifest/REPLAY')
        manifest = json.loads(package.read('EVIDENCE_MANIFEST.json'))
        if len(manifest['files']) != 1105 or len(manifest['cases']) != 26:
            raise ValueError('Archive evidence/case count differs')
        if set(manifest['files']) | {'EVIDENCE_MANIFEST.json', 'REPLAY.py'} != {v.filename for v in infos}:
            raise ValueError('Manifest does not enumerate the complete archive')
        target.mkdir()
        for item in infos:
            destination = target / item.filename
            destination.parent.mkdir(parents=True, exist_ok=True)
            result = hashlib.sha256()
            with package.open(item) as source, destination.open('xb') as output:
                for block in iter(lambda: source.read(1024 * 1024), b''):
                    result.update(block)
                    output.write(block)
            original_hashes[item.filename] = result.hexdigest()
        for name, expected in manifest['files'].items():
            if original_hashes[name] != expected:
                raise ValueError('Archived evidence hash differs: ' + name)
    return manifest, original_hashes


def main(args):
    started_epoch, started = time.time(), time.monotonic()
    deadline = started + args.seconds
    if args.output.exists():
        raise ValueError('Fresh verification output required')
    args.output.mkdir(parents=True)
    archive = args.archive.resolve()
    if archive.stat().st_size != EXPECTED_SIZE or digest(archive) != EXPECTED_SHA:
        raise ValueError('Final ZIP differs from the parent-frozen artifact')
    parents = []
    current = psutil.Process()
    for process in [current] + current.parents():
        try:
            parents.append(process_record(process))
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass
    controller = dict(source=str(Path(__file__).resolve()), sha256=digest(__file__),
        current_process=process_record(current), invoking_ancestry=parents,
        started_epoch=started_epoch, deadline_epoch=started_epoch + args.seconds,
        deadline_seconds=args.seconds, optimizer_calls=0)
    write(args.output / 'CONTROLLER.json', controller)
    manifest, member_hashes = safe_extract(archive, args.extract)
    source_hashes = {}
    for name, expected in manifest['files'].items():
        source = (REPO / name).resolve()
        if not source.is_relative_to(REPO) or digest(source) != expected:
            raise ValueError('Current original evidence differs from archive: ' + name)
        source_hashes[name] = expected
    package_root = REPO / 'validation/G1R19_python26_20261016'
    watched = {str(p.relative_to(REPO)).replace('\\', '/'): digest(p)
        for p in package_root.iterdir() if p.is_file() and
        (p.suffix in ('.py', '.csv') or p.name in ('CURRENT_RESULTS.json', 'CURRENT_FINDINGS_JA.md'))}
    before = dict(archive=dict(path=str(archive), size=archive.stat().st_size, sha256=EXPECTED_SHA),
        original_evidence_files=source_hashes, extracted_members=member_hashes, watched_root_files=watched)
    write(args.output / 'SHA_BEFORE.json', before)
    if inventory(args.extract) != member_hashes:
        raise ValueError('Extracted bytes or complete inventory differ before execution')
    environment = dict(os.environ, PYTHONDONTWRITEBYTECODE='1', PYTHONNOUSERSITE='1',
        OMP_NUM_THREADS='1', OPENBLAS_NUM_THREADS='1', MKL_NUM_THREADS='1',
        BLIS_NUM_THREADS='1', VECLIB_MAXIMUM_THREADS='1', NUMEXPR_NUM_THREADS='1')
    command = [str(args.python), '-X', 'utf8', str(args.extract / 'REPLAY.py'),
               '--output', str(args.output / 'replayed_cases')]
    processes = {}
    all_objects = {}
    aborted = False
    with (args.output / 'REPLAY.stdout.log').open('wb') as stdout, (args.output / 'REPLAY.stderr.log').open('wb') as stderr:
        child = subprocess.Popen(command, cwd=args.extract, env=environment, stdout=stdout, stderr=stderr)
        root_process = psutil.Process(child.pid)
        def capture():
            try:
                descendants = root_process.children(recursive=True)
            except psutil.NoSuchProcess:
                descendants = []
            for process in [root_process] + descendants:
                try:
                    record = process_record(process)
                    key = str(process.pid) + ':' + str(record['create_time'])
                    if key not in processes:
                        record.update(observed_first_epoch=time.time(), exited=False)
                        processes[key] = record
                        all_objects[key] = process
                except (psutil.NoSuchProcess, psutil.ZombieProcess):
                    pass
            for key, record in processes.items():
                if not record['exited'] and not unchanged_process(record):
                    record.update(exited=True, observed_exit_epoch=time.time(), os_returncode=None)
        previous_count = -1
        while child.poll() is None:
            capture()
            completed = len(list((args.output / 'replayed_cases').glob('*.json')))
            if completed != previous_count:
                print('REPLAY_PROGRESS', completed, '/26', flush=True)
                previous_count = completed
            write(args.output / 'PROCESS_LEDGER.json', dict(status='RUNNING', command=command, processes=processes))
            if time.monotonic() >= deadline:
                aborted = True
                # All descendants were observed under this owned root. Recheck
                # full command and birth immediately before child-first stop.
                for record in sorted(processes.values(), key=lambda v:v['create_time'], reverse=True):
                    if unchanged_process(record):
                        psutil.Process(record['pid']).kill()
                break
            time.sleep(.05)
        returncode = child.wait(timeout=15)
        capture()
        for record in processes.values():
            if unchanged_process(record):
                process = psutil.Process(record['pid'])
                try:
                    process.wait(timeout=5)
                except psutil.TimeoutExpired:
                    raise RuntimeError('Owned replay/auditor remained after parent exit')
            record['exited'] = not unchanged_process(record)
            record.setdefault('observed_exit_epoch', time.time())
            record.setdefault('os_returncode', None)
        for record in processes.values():
            if record['pid'] == child.pid:
                record['os_returncode'] = returncode
    after = dict(archive=dict(path=str(archive), size=archive.stat().st_size, sha256=digest(archive)),
        original_evidence_files={name:digest(REPO / name) for name in manifest['files']},
        extracted_members=inventory(args.extract),
        watched_root_files={name:digest(REPO / name) for name in watched})
    write(args.output / 'SHA_AFTER.json', after)
    cases = []
    for item in manifest['cases']:
        path = args.output / 'replayed_cases' / (item['case'] + '.json')
        result = json.loads(path.read_text(encoding='utf8')) if path.exists() else {}
        auditor_records = [p for p in processes.values() if '--case' in p['cmdline'] and
            p['cmdline'][p['cmdline'].index('--case')+1] == item['case']]
        cases.append(dict(case=item['case'], status=result.get('status'),
            target_met=result.get('target_met', False), root_complete=result.get('root_complete', False),
            relative_gap=result.get('relative_gap'), auditor_process_records=len(auditor_records),
            audited_result_sha256=digest(path) if path.exists() else None,
            original_cold_execution_verified=item['cold_execution_verified'],
            original_status=item['status']))
    all_exited = bool(processes) and all(p['exited'] for p in processes.values())
    passed = (not aborted and returncode == 0 and before == after and all_exited and
        len(cases) == 26 and all(c['target_met'] and c['root_complete'] and c['auditor_process_records'] for c in cases))
    write(args.output / 'PROCESS_LEDGER.json', dict(status='FINISHED' if all_exited else 'PROCESS_EXIT_UNCONFIRMED',
        command=command, replay_returncode=returncode, hard_deadline_exceeded=aborted, all_owned_exited=all_exited,
        note='Auditor exit codes are enforced by unmodified REPLAY subprocess.run(check=True); individual '
             'OS exit codes not captured are null rather than invented.', processes=processes))
    result = dict(status='PASS' if passed else 'FAIL', scope='Independent portable saved-field q8 pair/root replay; '
        'solver0; not a fresh cold26 execution', archive=str(archive), archive_sha256=EXPECTED_SHA,
        archive_bytes=EXPECTED_SIZE, original_evidence_files=1105, extracted_files=len(member_hashes),
        extraction_folder=str(args.extract), verification_folder=str(args.output),
        replay_returncode=returncode, external_deadline_seconds=args.seconds, hard_deadline_exceeded=aborted,
        all_owned_processes_exit_confirmed=all_exited, observed_process_count=len(processes),
        archive_original_and_extracted_sha_unchanged=before == after,
        verified_saved_pairs=sum(c['target_met'] and c['root_complete'] for c in cases),
        cold_normal_cases_unchanged=sum(c['original_cold_execution_verified'] for c in cases),
        optimizer_calls=0, elapsed_seconds=time.monotonic()-started, cases=cases,
        before_sha256=digest(args.output / 'SHA_BEFORE.json'), after_sha256=digest(args.output / 'SHA_AFTER.json'),
        controller_sha256=digest(__file__), started_epoch=started_epoch, finished_epoch=time.time())
    write(args.output / 'REPLAY_VERIFICATION.json', result)
    print(json.dumps({k:result[k] for k in ('status', 'verified_saved_pairs', 'cold_normal_cases_unchanged',
        'all_owned_processes_exit_confirmed', 'archive_original_and_extracted_sha_unchanged', 'elapsed_seconds')}), flush=True)
    return 0 if passed else 2


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--archive', type=Path, required=True)
    parser.add_argument('--extract', type=Path, required=True)
    parser.add_argument('--output', type=Path, required=True)
    parser.add_argument('--python', type=Path, required=True)
    parser.add_argument('--seconds', type=float, default=600.)
    args = parser.parse_args()
    if not 0 < args.seconds <= 600:
        raise ValueError('External guard must be positive and at most600s')
    raise SystemExit(main(args))
