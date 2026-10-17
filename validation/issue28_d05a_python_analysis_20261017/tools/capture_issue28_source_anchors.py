"""Snapshot exact Issue #28 review source anchors without running Excel or solvers."""
from __future__ import annotations
import argparse
import datetime as dt
import hashlib
import json
import subprocess
from pathlib import Path

REFERENCES = {
    'versions/G1R20_Fix4/src/vba/RPX_Fs.bas': [(248,277),(296,306),(889,917)],
    'versions/G1R20_Fix4/src/vba/RPX_AccuracyNative.bas': [(131,181),(263,285),(500,590)],
    'versions/G1R20_Fix4/src/vba/RPX_Conic.bas': [(412,446),(861,905),(968,1003)],
    'versions/G1R20_Fix4/src/vba/RPX_HSDCandidate.bas': [(104,110),(197,211),(255,267)],
    'versions/G1R20_Fix4/src/vba/RPX_Solve.bas': [(274,303),(1827,1831)],
    'versions/G1R20_Fix4/src/vba/RPX_Upper.bas': [(205,215)],
    'versions/G1R20_Fix4/src/vba/RPX_Audit.bas': [(289,298)],
    'versions/G1R19/native/unified/accuracy/src/lib.rs': [(108,124),(150,157)],
    'validation/G1R19_python26_20261016/python/experiment_core.py': [(169,224),(260,296)],
    'validation/G1R19_python26_20261016/python/original_dual.py': [(27,45),(49,79)],
    'validation/G1R19_python26_20261016/python/accuracy_driver.py': [(150,163)],
}

def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument('--repo', required=True, type=Path)
    ap.add_argument('--output', required=True, type=Path)
    args = ap.parse_args()
    repo = args.repo.resolve(strict=True)
    sources = []
    for path, ranges in REFERENCES.items():
        raw = (repo/path).read_bytes()
        lines = raw.decode('utf-8-sig').splitlines()
        excerpts = []
        for first, last in ranges:
            if first < 1 or last > len(lines) or last < first:
                raise ValueError(f'Invalid line range {path}:{first}-{last}')
            excerpts.append({'first_line': first, 'last_line': last,
                             'lines': [{'line': i, 'text': lines[i-1]} for i in range(first,last+1)]})
        sources.append({'path': path, 'bytes': len(raw),
                        'sha256': hashlib.sha256(raw).hexdigest(), 'anchors': excerpts})
    code = Path(__file__).read_bytes()
    report = Path(__file__).with_name('ISSUE28_CODE_FINDINGS.md')
    revision = subprocess.run(['git','rev-parse','HEAD'],cwd=repo,text=True,capture_output=True,check=True).stdout.strip()
    result = {'schema': 'RPFEM_ISSUE28_STATIC_SOURCE_ANCHORS_V1',
              'created_utc': dt.datetime.now(dt.timezone.utc).isoformat(), 'repo':str(repo),
              'git_head':revision, 'excel_runs':0, 'optimizer_calls':0,
              'scope':'Static source reading only; Issue-reported failure values are not reproduced measurements.',
              'sources':sources,
              'report':{'path':report.name,'bytes':report.stat().st_size,
                        'sha256':hashlib.sha256(report.read_bytes()).hexdigest()},
              'extractor':{'path':str(Path(__file__).resolve()),'bytes':len(code),
                           'sha256':hashlib.sha256(code).hexdigest()}}
    args.output.parent.mkdir(parents=True,exist_ok=True)
    args.output.write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    print(json.dumps({'output':str(args.output.resolve()),'source_files':len(sources),
                      'anchors':sum(len(s['anchors']) for s in sources), 'git_head':revision,
                      'report_sha256':result['report']['sha256']},ensure_ascii=False))

if __name__ == '__main__':
    main()
