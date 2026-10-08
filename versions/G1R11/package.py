from pathlib import Path
import csv
import hashlib
import json
import shutil
import zipfile

R = Path(__file__).resolve().parent

def read(relative):
    return json.loads((R / relative).read_text(encoding='utf-8'))

def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()

def write_csv(name, fields, rows):
    with (R / name).open('w', encoding='utf-8-sig', newline='') as stream:
        writer = csv.DictWriter(stream, fieldnames=fields)
        writer.writeheader()
        writer.writerows(rows)

release = read('Release_Verification.json')
book = Path(release['workbook'])
assert sha(book) == release['sha256']
identity = read('source_identity.json')
assert sha(Path(identity['path'])) == identity['sha256']
assert release['native_compile'] == release['saved_readback'] == 'PASS'
assert release['no_test_or_mock_code_in_workbook']
assert read('results/baseline_before.json')['baseline_intact']
assert read('results/baseline_after.json')['baseline_intact']
assert read('results/protected_originals.json')['original_workbooks_intact']

rows = []
def record(test, scope, evidence, detail=''):
    rows.append({'test': test, 'scope': scope, 'status': 'PASS',
                 'evidence': evidence, 'detail': detail})

static = read('results/static_invariants.json')
assert static['strict_cp932_crlf_roundtrip'] == 'PASS' and static['source_sha_intact']
record('不変式・変更範囲・配布文字コード', 'STATIC_SOURCE', 'results/static_invariants.json')
pretest = read('results/python_bracket_pretest.json')
assert pretest['random_cases_pass'] == 300
record('単調関数300例・探索復帰', pretest['scope'], 'results/python_bracket_pretest.json')
assert (R / 'results/staged_pretest.json').exists()
record('994/1000/1050要素の固定点・保存場監査',
       'PYTHON_FIXED_POINT_WITNESSES_NOT_FULL_ROOTS', 'results/staged_pretest.json',
       'NumericalError/AlmostSolvedはOPTIMALとして扱わない')
geometry = read('results/native_geometry.json')
assert geometry['native_python_exact_coordinate_triangles_materials'] == 'PASS'
assert geometry['legacy_geometry_matches_G1R9_bytes']
record('実VBAの新メッシュと従来分割同一性', 'NATIVE_EXCEL_GEOMETRY', 'results/native_geometry.json')
for file in ['native_controlled.json', 'native_brackets.json', 'native_settings_guard.json']:
    for item in read('results/' + file):
        record(item['test'], item.get('scope', 'NATIVE_INPUT_OR_CONTROL_TEST'),
               'results/' + file, item.get('result', ''))
for mode in ['lower', 'upper']:
    item = read('results/P06N_' + mode + '.json')
    assert item['result'].startswith('SOLVED;')
    record(item['test'], item['scope'], 'results/P06N_' + mode + '.json',
           item['result'] + f";seconds={item['seconds']:.6f}")
for item in read('results/native_independent_audits.json'):
    assert item['physical_audit']['pass']
    record(item['test'] + ' 原式監査', item['scope'], 'results/native_independent_audits.json')
    if 'unit_load_50digit' in item:
        assert item['unit_load_50digit']['numerical_audit_pass']
        record(item['test'] + ' 単位荷重50桁監査', 'INDEPENDENT_MPMATH_50_DIGITS',
               'results/native_independent_audits.json')
for item in read('results/native_small_roots.json'):
    assert item['result'].startswith('PASS;')
    record('小規模単層C5 Fs根 ' + item['policy'], item['scope'],
           'results/native_small_roots.json', item['result'])
assert read('results/single_material_exact_match.json')['fields_and_mesh_bytes_identical']
record('小規模単層のFs・求解数・保存場同一性', 'NATIVE_EXCEL_SMALL_ROOT_REGRESSION',
       'results/single_material_exact_match.json')
record('保存・再読込み・コンパイル・実装照合', 'NATIVE_EXCEL_RELEASE', 'Release_Verification.json')
record('旧baseline/reference・G1R9/G1R10原本保全', 'SHA256_PROTECTION',
       'results/protected_originals.json')
cases = [case for case in read('validation_runtime/data/cases.json') if case['id'] != 'D08A']
assert len(cases) == 25 and len({case['id'] for case in cases}) == 25
acceptance = []
for case in cases:
    item = {'test': case['id'] + ' 全Fs・適応回帰',
            'scope': 'NATIVE_EXCEL_FULL_CASE', 'status': 'NOT_RUN',
            'evidence': '', 'detail': '今回のG1R11全解析は未実施。間隙水圧対象外。'}
    rows.append(item)
    acceptance.append({'case': case['id'], 'status': 'NOT_RUN',
                       'lower_fs': '', 'upper_fs': '', 'gap': '',
                       'seconds': '', 'evidence': '', 'procedure': 'TEST_PROCEDURES.md'})
for test in ['G1R11 P06Nの全所要時間比較', 'GAP1%到達', '剛体あり通常SRM回帰']:
    rows.append({'test': test, 'scope': 'NATIVE_EXCEL_FULL_ANALYSIS', 'status': 'NOT_RUN',
                 'evidence': '', 'detail': '固定点・制御試験を全解析の代用にしない'})
write_csv('TEST_RESULTS.csv', ['test', 'scope', 'status', 'evidence', 'detail'], rows)
write_csv('excel_acceptance_g1r11.csv',
          ['case', 'status', 'lower_fs', 'upper_fs', 'gap', 'seconds', 'evidence', 'procedure'],
          acceptance)

fixes = [
    ('G11-01', '直接細分化が従来の良いメッシュ経路を飛ばす',
     'RPX_Adapt;RPX_Refine', '従来再配置とforceLegacy分割後に追加1回',
     'native_geometry.json;static_invariants.json;native_brackets.json'),
    ('G11-02', '1536要素まで成長し求解費用が増加',
     'RPX_Adapt', '追加予算min(1000,TARGET)、CYCLESと既存採否を維持',
     'native_controlled.json;native_geometry.json'),
    ('G11-03', '多層の下界探索で大きなFsの不要試行',
     'RPX_Fs', '監査済み同一モデル上界を試行位置の目安に使用。実符号を確認',
     'python_bracket_pretest.json;native_brackets.json'),
    ('G11-04', '改善策が単層や前段メッシュの求解を変える',
     'RPX_Fs', '使用材料2種類以上。staged前段800要素以下と剛体は対象外',
     'native_brackets.json;native_small_roots.json;single_material_exact_match.json'),
    ('G11-05', '追加段階で停止対象を通常復帰に混同する恐れ',
     'RPX_Adapt;RPX_Fs', '既知失敗のみ復元、致命的例外再送出、メモリ識別子保護',
     'native_controlled.json;native_brackets.json'),
    ('G11-06', '新設定の実行中変更・反映漏れ',
     '設定シート;既存RPX_Guard', '入力監視範囲内に追加。保存後再読込み照合',
     'native_settings_guard.json;Release_Verification.json'),
]
write_csv('ISSUE_FIX_MAP.csv', ['issue', 'problem', 'modules', 'fix', 'evidence'],
          [dict(zip(['issue', 'problem', 'modules', 'fix', 'evidence'], item)) for item in fixes])

evidence = R / 'prior_run_evidence'
evidence.mkdir(exist_ok=True)
old = R.parent / 'RPFEM_G1R10_Mesh_20261007'
for relative in ['FULL_RUN_REVIEW.md', 'evaluate_full_run.py', 'results/full_run_comparison.json']:
    shutil.copy2(old / relative, evidence / Path(relative).name)

desktop_book = R.parent / book.name
if desktop_book.exists():
    assert sha(desktop_book) == release['sha256'], 'ユーザーが変更した既存ファイルを保全する'
else:
    shutil.copy2(book, desktop_book)
assert sha(desktop_book) == release['sha256']

archive = R.parent / 'RPFEM_G1R11_Staged_20261007.zip'
if archive.exists():
    previous = read('Package_Verification.json')
    assert sha(archive) == previous['sha256'], '既存ZIPが変更されているため上書きしない'
excluded_names = {'Package_Verification.json', 'compile_failed_source.txt',
                  'native_small_roots_prototype.json'}
files = []
for path in R.rglob('*'):
    if not path.is_file() or '__pycache__' in path.parts or path.suffix == '.pyc':
        continue
    if path.name in excluded_names or (path.suffix == '.xlsm' and path != book):
        continue
    if path.name.startswith(('C5_LEGACY_', 'C5_AUDITED_UPPER_LIMIT_')):
        continue
    files.append(path)
with zipfile.ZipFile(archive, 'w', zipfile.ZIP_DEFLATED, compresslevel=6) as zipped:
    for path in sorted(files):
        zipped.write(path, path.relative_to(R))
with zipfile.ZipFile(archive) as zipped:
    assert zipped.testzip() is None
    assert hashlib.sha256(zipped.read(book.name)).hexdigest() == release['sha256']
    assert not any(name.startswith('tests/') and name.endswith('.xlsm') for name in zipped.namelist())
result = {'archive': str(archive), 'sha256': sha(archive), 'bytes': archive.stat().st_size,
          'members': len(files), 'CRC': 'PASS', 'workbook_member_sha256': 'PASS',
          'desktop_workbook': str(desktop_book), 'workbook_sha256': sha(desktop_book),
          'test_rows_pass': sum(row['status'] == 'PASS' for row in rows),
          'test_rows_not_run': sum(row['status'] == 'NOT_RUN' for row in rows),
          'full_case_acceptance_rows_not_run': len(acceptance)}
assert sha(Path(identity['path'])) == identity['sha256']
(R / 'Package_Verification.json').write_text(json.dumps(result, ensure_ascii=False, indent=2),
                                           encoding='utf-8')
print(result)
