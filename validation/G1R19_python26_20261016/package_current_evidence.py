"""Archive selected verified pairs and their exact code/input bytes.

Partial bundles require an explicit flag and remain labelled PARTIAL. This is
a saved-field replay bundle, not a claim that all cold executions completed.
"""
from datetime import datetime, timezone
from pathlib import Path
import argparse
import hashlib
import json
import zipfile

import collect_current as collector

REPO = collector.REPO


def digest(data):
    return hashlib.sha256(data).hexdigest()


REPLAY = '''from pathlib import Path
import argparse, hashlib, json, subprocess, sys
root=Path(__file__).resolve().parent
p=argparse.ArgumentParser();p.add_argument('--output',type=Path,required=True);args=p.parse_args()
manifest=json.loads((root/'EVIDENCE_MANIFEST.json').read_text(encoding='utf8'))
for name,expected in manifest['files'].items():
    if hashlib.sha256((root/name).read_bytes()).hexdigest()!=expected:
        raise SystemExit('Evidence bytes changed: '+name)
args.output.mkdir(parents=True,exist_ok=False)
for case in manifest['cases']:
    out=(args.output/(case['case']+'.json')).resolve()
    command=[sys.executable,'-X','utf8',str(root/case['auditor']),
             '--case',case['case'],'--folder',str(root/case['folder']),
             '--target-gap','.02','--output',str(out)]
    subprocess.run(command,cwd=root,check=True,timeout=180)
    result=json.loads(out.read_text(encoding='utf8'))
    if not result.get('target_met') or not result.get('root_complete'):
        raise SystemExit('Replay failed: '+case['case'])
print('Verified saved-field pairs:',len(manifest['cases']),
      '; no new optimization; cold execution statuses remain in the manifest')
'''


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument('--report', type=Path, default=collector.ROOT/'CURRENT_RESULTS.json')
    ap.add_argument('--output', type=Path, required=True)
    ap.add_argument('--allow-partial', action='store_true')
    args = ap.parse_args()
    if args.output.exists():
        raise FileExistsError('Refusing to replace an evidence archive')
    report = collector.read(args.report)
    rows = [row for row in report['cases'] if row.get('scientific_pair_verified')]
    if len(rows) != report['expected_cases'] and not args.allow_partial:
        raise ValueError('All cases have not passed; an interim bundle needs --allow-partial')
    files = {}

    def add(name, expected=None):
        path = Path(name)
        path = (path if path.is_absolute() else REPO/path).resolve()
        if not path.is_relative_to(REPO.resolve()):
            raise ValueError('Evidence file is outside this repository')
        data = path.read_bytes(); actual = digest(data)
        if expected is not None and actual != expected:
            raise ValueError('Evidence changed: '+str(path))
        key = path.relative_to(REPO).as_posix()
        if key in files and files[key] != (data, actual):
            raise ValueError('Conflicting evidence versions: '+key)
        files[key] = (data, actual)

    cases = []
    for row in rows:
        case = row['case']; audit_path = REPO/row['audit']; summary_path = REPO/row['summary']
        geometry_replay = row.get('kind') == 'known_recipe_geometry_replay'
        verified = collector.validate(audit_path, summary_path, case, geometry_replay=geometry_replay)
        audit, summary = collector.read(audit_path), collector.read(summary_path)
        add(audit_path); add(summary_path)
        if geometry_replay:
            manifest_path = summary_path.parent/summary['component_manifest']['path']
            add(manifest_path, summary['component_manifest']['sha256'])
            manifest = collector.read(manifest_path)
            for component in manifest['components']:
                for key in ('ownership_file', 'summary', 'execution_record',
                            'independent_primal_audit', 'independent_dual_audit'):
                    ref = component.get(key)
                    if ref is not None:
                        add(ref['path'], ref['sha256'])
                        if key == 'independent_dual_audit':
                            replay = collector.read(Path(ref['path']))
                            for map_key in ('code_and_input_sha256_before', 'code_and_input_sha256_after'):
                                for name, expected in replay.get(map_key, {}).items():
                                    add(name, expected)
                for key in ('component_start_source_sha256', 'component_end_source_sha256', 'input_provenance'):
                    for name, expected in component[key].items():
                        add(name, expected)
                for name, expected in component['component_artifacts_sha256'].items():
                    add(Path(component['folder'])/name, expected)
            for name, expected in manifest['collector_reaudit_source_sha256'].items():
                add(name, expected)
            for key in ('collector_script',):
                ref = manifest.get(key)
                if ref is not None:
                    add(ref['path'], ref['sha256'])
            for key in ('prior_timing_file',):
                ref = manifest['timing'].get(key)
                if ref is not None:
                    add(ref['path'], ref['sha256'])
            # A posthoc pair has several genuine producers, not one invented
            # environment/start map. Its independent replay uses the base q8 auditor.
            producer = collector.ROOT
            base_manifest = collector.read(producer/'INPUTS_26.json')
            base_entry = next(item for item in base_manifest['cases'] if item['case'] == case)
            add(producer/'INPUTS_26.json')
            add(base_entry['input']['path'])
            if 'initial_mesh' in base_entry:
                add(base_entry['initial_mesh']['path'], base_entry['initial_mesh']['sha256'])
        else:
            manifest_path = summary_path.parent/'source_manifest.json'
            add(manifest_path)
            maps = [collector.read(manifest_path)['source_sha256'], summary.get('wrapper_source_sha256', {})]
            for mapping in maps:
                for name, expected in mapping.items():
                    add(name, expected)
            producer = REPO/verified['provenance']['producer_root']
            manifest = collector.read(producer/'INPUTS_26.json')
            entry = next(item for item in manifest['cases'] if item['case'] == case)
            add(producer/'INPUTS_26.json')
            add(entry['input']['path'])
            if 'initial_mesh' in entry:
                add(entry['initial_mesh']['path'], entry['initial_mesh']['sha256'])
        for side in audit['sides'].values():
            for name, expected in side.get('tracked_file_sha256', side.get('file_sha256', {})).items():
                add(name, expected)
        if audit.get('load_model_consistency'):
            consistency = audit['load_model_consistency']
            add(consistency['model_verification_path'], consistency['model_verification_sha256'])
        auditor = producer/'audit_case.py'
        if not auditor.is_file():
            auditor = collector.ROOT/'audit_case.py'
        add(auditor)
        # Dry retries reuse the base independent auditor and its original
        # builder. The wet retry carries its distinct common-load auditor.
        if auditor.parent == collector.ROOT:
            for path in (collector.ROOT/'python').rglob('*.py'):
                add(path)
        cases.append(dict(case=case, folder=summary_path.parent.relative_to(REPO).as_posix(),
                          auditor=auditor.relative_to(REPO).as_posix(),
                          kind=row.get('kind'), status=row['status'],
                          scientific_pair_verified=True,
                          cold_execution_verified=row.get('cold_execution_verified', False),
                          gap_percent=verified['gap_percent']))
    add(args.report)
    for path in (Path(__file__), collector.ROOT/'collect_current.py', collector.ROOT/'evidence_provenance.py',
                 collector.ROOT/'composite_provenance.py'):
        add(path)
    manifest = dict(status='ALL_NUMERICAL_PAIRS' if len(rows) == report['expected_cases'] else 'PARTIAL',
                    generated_utc=datetime.now(timezone.utc).isoformat(),
                    scope='Exact saved-pair replay; no new solves and no cold-completion claim',
                    expected_cases=report['expected_cases'], verified_cases=len(rows), cases=cases,
                    historical_producer_paths='Retained unchanged in original JSON; replay uses the portable relative folders above',
                    files={name: value[1] for name, value in sorted(files.items())})
    args.output.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(args.output, 'x', compression=zipfile.ZIP_DEFLATED, compresslevel=6) as archive:
        for name, (data, _) in sorted(files.items()):
            archive.writestr(name, data)
        archive.writestr('EVIDENCE_MANIFEST.json', json.dumps(manifest, indent=2, ensure_ascii=False)+'\n')
        archive.writestr('REPLAY.py', REPLAY)
    with zipfile.ZipFile(args.output) as archive:
        if archive.testzip() is not None:
            raise ValueError('Archive CRC verification failed')
        for name, (_, expected) in files.items():
            if digest(archive.read(name)) != expected:
                raise ValueError('Archived bytes differ: '+name)
    record = dict(status=manifest['status'], verified_cases=len(rows), file_count=len(files),
                  archive=args.output.name, bytes=args.output.stat().st_size,
                  sha256=digest(args.output.read_bytes()), crc_and_all_member_sha256='PASS',
                  saved_field_replay_not_yet_run=True)
    args.output.with_suffix('.json').write_text(json.dumps(record, indent=2)+'\n', encoding='utf8')
    print(json.dumps(record))


if __name__ == '__main__':
    main()
