"""Package completed, verified native evidence; exclude scheduler placeholders."""
from pathlib import Path
import json,shutil,zipfile,hashlib
from prepare_inputs import ROOT,REPO,SOURCE,SOURCE_SHA
from case_locations import case_folder,evidence_root

def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()

def zipped(path,folder):
    with zipfile.ZipFile(path,'w',compression=zipfile.ZIP_DEFLATED,compresslevel=9) as z:
        for p in sorted(folder.rglob('*')):
            if p.is_file():z.write(p,p.relative_to(folder))
    with zipfile.ZipFile(path) as z:assert z.testzip() is None

def run():
    summary=json.loads((ROOT/'results/final_summary.json').read_text(encoding='utf-8'))
    assert summary['all25_native_executed'] and len(summary['results'])==25
    assert sha(SOURCE)==SOURCE_SHA
    delivery=ROOT/'delivery';assert not delivery.exists();delivery.mkdir()
    evidence=delivery/'evidence';books=delivery/'workbooks';evidence.mkdir();books.mkdir()
    for name in ['FINAL_REPORT.md','FINAL_RESULTS.csv','source_identity.json','environment.json','Test_Case_Catalog.md']:
        shutil.copy2(ROOT/name,evidence/name)
    for name in ['FINAL_REPORT.md','FINAL_RESULTS.csv']:shutil.copy2(ROOT/name,books/name)
    shutil.copytree(ROOT/'inputs',evidence/'inputs')
    for p in ROOT.glob('*.py'):shutil.copy2(p,evidence/p.name)
    shutil.copytree(ROOT/'tests',evidence/'tests',ignore=shutil.ignore_patterns('__pycache__'))
    (evidence/'results').mkdir()
    for name in ['final_summary.json','template_check.json','solve_costs.json','solve_costs.csv','native25_terminal.json','invariants.json','pure_c0_root_review.json','preservation_checks.json']:
        p=ROOT/'results'/name
        if p.exists():shutil.copy2(p,evidence/'results'/name)
    original=REPO/'versions/G1R12';runtime=evidence/'runtime/versions/G1R12';runtime.mkdir(parents=True)
    shutil.copy2(SOURCE,runtime/SOURCE.name)
    shutil.copytree(original/'src',runtime/'src')
    shutil.copytree(original/'validation_runtime',runtime/'validation_runtime',ignore=shutil.ignore_patterns('__pycache__'))
    (runtime/'tests').mkdir();shutil.copy2(original/'tests/native_helpers.py',runtime/'tests/native_helpers.py')
    (runtime/'results').mkdir()
    manifest=[]
    for row in summary['results']:
        folder=case_folder(row['case']);out=evidence/'cases'/row['case'];out.mkdir(parents=True)
        for p in sorted(folder.rglob('*')):
            if not p.is_file() or p.suffix in ['.xlsm','.tmp']:continue
            dest=out/p.relative_to(folder);dest.parent.mkdir(parents=True,exist_ok=True);shutil.copy2(p,dest)
        (out/'derived_result.json').write_text(json.dumps(row,ensure_ascii=False,indent=2),encoding='utf-8')
        if row.get('workbook'):
            actual=evidence_root(row['case'])/row['workbook'];assert sha(actual)==row['workbook_sha256']
            dst=books/'cases'/row['case']/actual.name;dst.parent.mkdir(parents=True);shutil.copy2(actual,dst)
            manifest.append({'case':row['case'],'path':str(dst.relative_to(books)),'sha256':sha(dst),'status':row['status']})
    diag=evidence/'surcharge_diagnosis';diag.mkdir()
    for p in (ROOT/'surcharge_diagnosis').iterdir():
        if p.is_file() and p.suffix!='.xlsm':shutil.copy2(p,diag/p.name)
    shutil.copytree(ROOT/'underflow_diagnosis',evidence/'underflow_diagnosis')
    shutil.copytree(ROOT/'pure_c0_feasibility',evidence/'pure_c0_feasibility')
    shutil.copytree(ROOT/'pure_c0_upper_selection',evidence/'pure_c0_upper_selection')
    meshdiag=ROOT/'mesh_identity_target538'
    (evidence/'mesh_identity_target538').mkdir()
    for name in ['comparison.json','production.bin','catalog_adapter.bin']:shutil.copy2(meshdiag/name,evidence/'mesh_identity_target538'/name)
    smoke=evidence/'harness_validation';smoke.mkdir()
    for p in (ROOT/'harness_validation').rglob('*'):
        if p.is_file() and p.suffix in ['.json','.bin']:
            dst=smoke/p.relative_to(ROOT/'harness_validation');dst.parent.mkdir(parents=True,exist_ok=True);shutil.copy2(p,dst)
    repro='''# 再実行の手順

通常のExcel解析は、ケースのXLSMを開いて入力と設定を確認し「解析」を実行します。PythonやVBProjectアクセスは通常の解析には不要です。保存結果の再計算では実行時間が再びかかります。

このバッチ検証を再現する場合は、証拠フォルダーそのものを実行先にせず、新しい空フォルダーへルートのPythonスクリプト、tests、runtime、Test_Case_Catalog.mdだけをコピーしてください。結果ケースやprepared_template.xlsmをコピーすると上書き防止で停止します。

環境はenvironment.jsonに記録しています。Windowsの64bit Excel、Python、numpy・scipy・clarabel・mpmath・pywin32・oletools・matplotlibを使用しています。テンプレートの証拠モジュール作成には許可済みのVBProjectアクセスが必要です。スクリプトはExcelの信頼・マクロ・レジストリ設定を変更しません。

PowerShellで新しいフォルダーへ移動し、Python実行環境に合わせて次の操作を行います。

```powershell
$env:RPFEM_REPO = Join-Path (Get-Location) 'runtime'
$env:RPFEM_CATALOG = Join-Path (Get-Location) 'Test_Case_Catalog.md'
python -X utf8 prepare_inputs.py
python -X utf8 run_native25.py prepare batch
python -X utf8 watch_completed.py
python -X utf8 summarize_log_cost.py
python -X utf8 review_pure_c0_roots.py
python -X utf8 compare_invariants.py
python -X utf8 finalize_report.py
```

この手順は新しいケースコピーで25件を実行します。SOURCE_SHAが違えば停止します。watch_completed.pyは実解析終了後に保存ブックのVBA・出力を照合します。続いて求解時間、純c=0の根証明の記録、平行移動などの比較を集計し、finalize_report.pyで全25件の最終結果を確定します。バッチの待ち行列移管continue_queue.pyは今回の実行を高速化する内部手順の記録で、標準の再実行には必要ありません。監査・保存表・VBA照合はaudit_batch.py、check_saved_output.py、check_saved_sources.py、集計定義はgrade_results.pyで確認できます。

修復候補の診断フォルダーは今回の保存済み証拠です。このバッチ手順だけでは診断の追加実験を再作成しません。再生成した報告書の診断リンクを読む場合は、証拠ZIPの対応する診断フォルダーも参照してください。

証拠ZIPではケースのログ・場・JSONをcases/<ID>へまとめ、XLSMは別のworkbooks ZIPへ分けています。診断用の小規模試験・メッシュ生成試験・Python修復候補を25件の実Excel合格に加算していません。D07A、D03Nの修復対策はVBA未実装です。
'''
    (evidence/'REPRODUCE_JA.md').write_text(repro,encoding='utf-8')
    (books/'WORKBOOK_MANIFEST.json').write_text(json.dumps(manifest,ensure_ascii=False,indent=2),encoding='utf-8')
    for package in [evidence,books]:
        entries=[{'path':str(p.relative_to(package)),'bytes':p.stat().st_size,'sha256':sha(p)} for p in sorted(package.rglob('*')) if p.is_file()]
        (package/'FILE_MANIFEST.json').write_text(json.dumps(entries,ensure_ascii=False,indent=2),encoding='utf-8')
    files=[]
    for folder,name in [(evidence,'RPFEM_G1R12_Native25_Evidence_20261009.zip'),(books,'RPFEM_G1R12_Native25_Workbooks_20261009.zip')]:
        path=delivery/name;zipped(path,folder);files.append({'name':name,'bytes':path.stat().st_size,'sha256':sha(path)})
    (delivery/'SHA256SUMS.txt').write_text(''.join(r['sha256']+'  '+r['name']+'\n' for r in files),encoding='utf-8')
    (delivery/'package_summary.json').write_text(json.dumps(files,indent=2),encoding='utf-8')
    print(json.dumps(files,ensure_ascii=False))

if __name__=='__main__':run()
