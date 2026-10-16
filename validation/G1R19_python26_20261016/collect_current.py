"""Keep every cold attempt, and report independently verified 2% pairs only.

This collector does not solve or import historical witnesses. A failed first
attempt remains visible when a separately versioned cold retry succeeds.
"""
from pathlib import Path
from decimal import Decimal, ROUND_HALF_UP
import argparse
import csv
import hashlib
import json
import math

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[1]


def read(path):
    return json.loads(path.read_text(encoding='utf8'))


def relative(path):
    return path.resolve().relative_to(REPO.resolve()).as_posix()


def validate(audit_path, summary_path, case):
    audit, summary = read(audit_path), read(summary_path)
    if audit.get('case') != case or summary.get('case') != case:
        raise ValueError('Case identity differs')
    if not audit.get('target_met') or not audit.get('root_complete'):
        raise ValueError('Independent audit does not certify the target and both roots')
    if not summary.get('target_met') or not summary.get('source_unchanged'):
        raise ValueError('The producer did not finish with unchanged sources and the target')
    lower, upper = audit['lower_fs'], audit['upper_fs']
    if not all(math.isfinite(v) for v in (lower, upper)) or not 0 < lower <= upper:
        raise ValueError('Nonfinite, reversed or nonpositive bounds')
    gap = 2 * (upper - lower) / (upper + lower)
    if gap > .02 or abs(gap-audit['relative_gap']) > 1e-14:
        raise ValueError('Recomputed symmetric gap fails')
    models = []
    for mode in ('lower', 'upper'):
        side = audit['sides'][mode]
        physical = side.get('audit', side)
        if not side.get('root_complete') or not physical.get('physical_pass') or not physical.get('witness'):
            raise ValueError('An original physical gate or own-root check failed')
        if side['fs'] != (lower if mode == 'lower' else upper):
            raise ValueError('Audited endpoint differs')
        models.append(side['physical_model_id'])
        tracked = side.get('tracked_file_sha256', side.get('file_sha256', {}))
        if not tracked:
            raise ValueError('Independent audit did not track its input bytes')
        for name, expected in tracked.items():
            path = Path(name).resolve()
            if not path.is_relative_to(REPO.resolve()):
                raise ValueError('Evidence file is outside this repository')
            if hashlib.sha256(path.read_bytes()).hexdigest() != expected:
                raise ValueError('An audited file changed: '+relative(path))
    if models[0] != models[1]:
        raise ValueError('Lower and upper physical models differ')
    source = read(REPO/'docs/benchmarks/G1R17_native26_20261013/data/inputs'/f'{case}.json')
    if source.get('source_case', {}).get('pore'):
        # A matching nominal ID alone cannot establish a common continuum
        # load when nodal P1 pore-pressure gradients change with the mesh.
        consistency = audit.get('load_model_consistency', {})
        if consistency.get('verified') is not True or consistency.get('physical_model_id') != models[0]:
            raise ValueError('A mesh-independent common pore-pressure load has not been independently verified')
        verification = (REPO/consistency['model_verification_path']).resolve()
        if not verification.is_relative_to(REPO.resolve()):
            raise ValueError('Common-load verification is outside this repository')
        if hashlib.sha256(verification.read_bytes()).hexdigest() != consistency['model_verification_sha256']:
            raise ValueError('Common-load verification bytes changed')
        if read(verification).get('status') != 'PASS':
            raise ValueError('Common-load geometry verification failed')
    rounded = [str(Decimal(str(v)).quantize(Decimal('.01'), rounding=ROUND_HALF_UP)) for v in (lower, upper)]
    return dict(lower_fs=lower, upper_fs=upper, gap_percent=gap*100,
                lower_elements=audit['sides']['lower']['elements'],
                upper_elements=audit['sides']['upper']['elements'],
                physical_model_id=models[0], lower_rounded_2dp=rounded[0],
                upper_rounded_2dp=rounded[1], rounded_2dp_agree=rounded[0] == rounded[1],
                audit=relative(audit_path), summary=relative(summary_path))


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument('--batch', type=Path, required=True)
    ap.add_argument('--additional-batch', type=Path, action='append', default=[])
    ap.add_argument('--retries', type=Path, default=ROOT/'COLD_RETRIES.json')
    ap.add_argument('--output', type=Path, default=ROOT)
    args = ap.parse_args()
    progress = read(args.batch/'PROGRESS.json')
    rows = {case: dict(case=case, status='PENDING', verified=False, attempts=[])
            for case in progress['provenance']['expected_cases']}
    attempts = []
    batches = [(args.batch, 'v1', progress)] + [
        (p, 'v1_'+p.name, read(p/'PROGRESS.json')) for p in args.additional_batch]
    for batch, version, batch_progress in batches:
        for record in batch_progress['cases']:
            case = record['case']
            if case not in rows:
                raise ValueError('An additional batch contains a case outside the 26-case catalog')
            attempt = dict(case=case, version=version, status=record['status'], seconds=record.get('seconds'),
                           kind='fresh_cold', selected=False, process_exit_code=record.get('returncode'),
                           lower_fs=record.get('lower_fs'), upper_fs=record.get('upper_fs'),
                           gap_percent=record.get('gap_percent'), batch=relative(batch))
            if record['status'] == 'PASS' and record.get('returncode') == 0:
                try:
                    attempt.update(validate(batch/f'{case}_independent_audit.json',
                                            batch/'cases'/case/'summary.json', case), status='PASS', verified=True)
                except (KeyError, ValueError, OSError) as error:
                    attempt.update(status='EVIDENCE_REJECTED', verified=False, error=str(error))
            else:
                attempt['verified'] = False
            attempts.append(attempt)
    if args.retries.exists():
        for record in read(args.retries):
            attempt = dict(record, kind='fresh_cold', selected=False, verified=False)
            try:
                attempt.update(validate(REPO/record['audit'], REPO/record['summary'], record['case']),
                               status='PASS', verified=True)
            except (KeyError, ValueError, OSError) as error:
                attempt.update(status='EVIDENCE_REJECTED', error=str(error))
            attempts.append(attempt)
    for attempt in attempts:
        row = rows[attempt['case']]
        row['attempts'].append(attempt)
        if attempt['verified'] or not row['verified']:
            row.update({k:v for k,v in attempt.items() if k not in ('attempts', 'selected')})
    for row in rows.values():
        for attempt in reversed(row['attempts']):
            if attempt['version'] == row.get('version'):
                attempt['selected'] = True
                break
    count = sum(r['verified'] for r in rows.values())
    report = dict(scope='All 26 catalog cases; fresh cold attempts only; no historical replay counted',
                  verified_cases=count, expected_cases=len(rows), all_26_verified=count == len(rows),
                  target='2*(upper-lower)/(upper+lower) <= 0.02 plus original q8 physical gates and own roots',
                  rounding_note='ROUND_HALF_UP at two decimal places is a separate display check, not a stop target',
                  cases=list(rows.values()))
    args.output.mkdir(parents=True, exist_ok=True)
    (args.output/'CURRENT_RESULTS.json').write_text(json.dumps(report, indent=2, ensure_ascii=False)+'\n', encoding='utf8')
    fields = ['case','status','verified','version','seconds','lower_fs','upper_fs','gap_percent',
              'lower_elements','upper_elements','lower_rounded_2dp','upper_rounded_2dp','rounded_2dp_agree','audit']
    with (args.output/'CURRENT_RESULTS.csv').open('w', encoding='utf8', newline='') as stream:
        writer = csv.DictWriter(stream, fieldnames=fields, extrasaction='ignore')
        writer.writeheader(); writer.writerows(rows.values())
    print(f'{count}/{len(rows)} fresh cold pairs independently verified at GAP <= 2%')
    for row in rows.values():
        if row['status'] != 'PENDING':
            print(row['case'], row['status'], row.get('version'), row.get('gap_percent'))


if __name__ == '__main__':
    main()
