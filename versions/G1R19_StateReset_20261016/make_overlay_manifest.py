"""Create/check four explicit-call changes in this isolated, unpublished overlay."""
from pathlib import Path
import argparse
import hashlib
import json

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[1]
EDITS = {
    'RPX_Fs.bas': ('versions/G1R19_2pct/src/vba/RPX_Fs.bas',
                   '            RPX_ResetPreparedState: RPX_Compile: prepared = True',
                   '            Call RPX_ResetPreparedState: RPX_Compile: prepared = True'),
    'RPX_Conic.bas': ('versions/G1R19/src/vba/RPX_Conic.bas',
                      '    RPX_OriginalGuideClear: OriginalFarkasReason = "scope"',
                      '    Call RPX_OriginalGuideClear: OriginalFarkasReason = "scope"'),
    'RPX_Solve.bas': ('versions/G1R19/src/vba/RPX_Solve.bas',
                      '    HContext: HEnsureDD XN + XM',
                      '    Call HContext: HEnsureDD XN + XM'),
    'RPX_Seep.bas': ('versions/G1R19/src/vba/RPX_Seep.bas',
                     'RPX_SeepDefaults:     SRelax = relax',
                     'Call RPX_SeepDefaults:     SRelax = relax'),
}


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def normalize(text):
    return text.replace('\r\n', '\n').replace('\r', '\n')


def expected_text(name):
    relative, before, after = EDITS[name]
    baseline = normalize((REPO/relative).read_text(encoding='utf8'))
    lines = baseline.splitlines()
    if lines.count(before) != 1:
        raise ValueError('Baseline anchor is not unique: '+name)
    return baseline.replace(before+'\n', after+'\n')


def manifest_data():
    records = []
    for name, (relative, before, after) in EDITS.items():
        baseline, source, imported = REPO/relative, ROOT/'src/vba'/name, ROOT/'import'/name
        expected = expected_text(name)
        if normalize(source.read_text(encoding='utf8')) != expected:
            raise ValueError('More than the permitted single-line change: '+name)
        if source.read_bytes().decode('utf8') != imported.read_bytes().decode('cp932'):
            raise ValueError('Source/import text differs: '+name)
        if source.read_bytes() != expected.replace('\n', '\r\n').encode('utf8'):
            raise ValueError('UTF8 source must have the exact expected CRLF bytes: '+name)
        line = normalize(baseline.read_text(encoding='utf8')).splitlines().index(before)+1
        records.append(dict(module=name.removesuffix('.bas'), baseline=relative,
                            baseline_sha256=sha(baseline), source='src/vba/'+name,
                            source_sha256=sha(source), imported='import/'+name,
                            import_sha256=sha(imported), changed_line=line,
                            before=before, after=after, decoded_text_equal=True,
                            normalized_diff='one line; insert Call plus one space only'))
    tracked = ['README_JA.md', 'tests/test_contract.py', 'tests/make_probe_sources.py',
               'tests/MICROTEST_JA.md']
    return dict(schema='rpfem.statereset.overlay.v1',
                status='STATIC_CONTRACT_ONLY; EXCEL_COMPILE_AND_MICRO_PENDING',
                scope='Four explicit parameterless Sub calls; no established false-PASS claim',
                production_artifacts_modified=False, excel_runs=0, solver_calls=0,
                base_composition='G1R19 plus authoritative G1R19_2pct overlay; then these four modules',
                unchanged_contract=['physical model', 'objective', 'q', 'raw physical gates',
                                    'root signs and normalized tolerance', 'GAP target', 'DLL'],
                modules=records,
                references=[dict(path=p, sha256=sha(ROOT/p)) for p in tracked])


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--initialize', action='store_true', help='Create only missing source/import files')
    parser.add_argument('--check', action='store_true', help='Check existing manifest without writing it')
    args = parser.parse_args()
    if args.initialize:
        for name in EDITS:
            text = expected_text(name).replace('\n', '\r\n')
            for directory, encoding in [('src/vba','utf8'), ('import','cp932')]:
                path = ROOT/directory/name
                path.parent.mkdir(parents=True, exist_ok=True)
                if not path.exists():
                    path.write_bytes(text.encode(encoding))
    data = manifest_data()
    manifest = ROOT/'SOURCE_MANIFEST.json'
    if args.check:
        if json.loads(manifest.read_text(encoding='utf8')) != data:
            raise ValueError('Manifest or referenced bytes changed')
    else:
        manifest.write_text(json.dumps(data, ensure_ascii=False, indent=2)+'\n', encoding='utf8')
    print(json.dumps(dict(status=data['status'], source_import_pairs=len(data['modules']),
                          manifest_sha256=sha(manifest)), ensure_ascii=False))


if __name__ == '__main__':
    main()
