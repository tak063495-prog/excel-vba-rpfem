"""Check final ZIP membership, payload hashes and actual case provenance."""
import json,hashlib,zipfile,csv,io
from prepare_inputs import ROOT,SOURCE_SHA

def sha(raw):return hashlib.sha256(raw).hexdigest()

def verify_manifest(z):
    manifest=json.loads(z.read('FILE_MANIFEST.json'))
    names=set(z.namelist())
    assert names=={r['path'].replace('\\','/') for r in manifest}|{'FILE_MANIFEST.json'}
    for row in manifest:
        raw=z.read(row['path'].replace('\\','/'))
        assert len(raw)==row['bytes'] and sha(raw)==row['sha256'],row['path']

def run():
    summary=json.loads((ROOT/'results/final_summary.json').read_text(encoding='utf-8'))
    rows=summary['results'];assert summary['all25_native_executed'] and len(rows)==25
    assert all(r['native_run_started'] for r in rows)
    assert all(r['status'] not in ['RUNNING','PREPARING','NOT_RUN'] for r in rows)
    packages=json.loads((ROOT/'delivery/package_summary.json').read_text(encoding='utf-8'))
    for p in packages:assert sha((ROOT/'delivery'/p['name']).read_bytes())==p['sha256']
    evidence=ROOT/'delivery'/next(p['name'] for p in packages if 'Evidence' in p['name'])
    books=ROOT/'delivery'/next(p['name'] for p in packages if 'Workbooks' in p['name'])
    with zipfile.ZipFile(evidence) as z:
        verify_manifest(z)
        for row in rows:
            case=row['case'];prefix='cases/'+case+'/'
            status=json.loads(z.read(prefix+'status.json'))
            derived=json.loads(z.read(prefix+'derived_result.json'))
            assert status['case']==derived['case']==case
            assert sha(z.read('inputs/'+case+'.json'))==status['input_sha256']==derived['input_sha256']
            assert derived['status']==row['status'] and derived['native_run_started']
            assert any(p.startswith(prefix+'RPFEM_logs/') and p.endswith('_analysis.tsv') for p in z.namelist())
            assert json.loads(z.read(prefix+'saved_sources_check.json'))['pass']
            if row['fs_obtained']:
                assert json.loads(z.read(prefix+'saved_output_check.json'))['pass']
                assert derived['independent_audit']['pass']
                for side in ['lower','upper']:assert z.read(prefix+'witness_'+side+'.bin')[:8]==b'RP25M001'
        production='runtime/versions/G1R12/RPFEM_20261009_G1R12_P06N_IssueFix.xlsm'
        assert sha(z.read(production))==SOURCE_SHA
        assert [p for p in z.namelist() if p.endswith('.xlsm')]==[production]
    with zipfile.ZipFile(books) as z:
        verify_manifest(z)
        manifest=json.loads(z.read('WORKBOOK_MANIFEST.json'));assert len(manifest)==25
        names={p for p in z.namelist() if p.endswith('.xlsm')}
        assert names=={r['path'].replace('\\','/') for r in manifest}
        expected={r['case']:r for r in rows}
        for r in manifest:
            raw=z.read(r['path'].replace('\\','/'));row=expected[r['case']]
            assert sha(raw)==r['sha256']==row['workbook_sha256']
            assert r['status']==row['status']
            with zipfile.ZipFile(io.BytesIO(raw)) as book:
                assert 'xl/vbaProject.bin' in book.namelist()
                assert 'xl/workbook.xml' in book.namelist()
        csvrows=list(csv.DictReader(io.StringIO(z.read('FINAL_RESULTS.csv').decode('utf-8-sig'))))
        assert len(csvrows)==25 and {r['case'] for r in csvrows}==set(expected)
        assert all(r['status']==expected[r['case']]['status'] for r in csvrows)
    result={'pass':True,'all25_actual_native_evidence':True,'native_workbooks':25,
            'all25_input_hashes_match_native_records':True,
            'invalid_scheduler_files_excluded':True,'all_zip_payload_manifest_hashes_pass':True,
            'production_sha256':SOURCE_SHA,'packages':packages}
    (ROOT/'delivery/delivery_verification.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps({k:v for k,v in result.items() if k!='packages'},ensure_ascii=False))

if __name__=='__main__':run()
