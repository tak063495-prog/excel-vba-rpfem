"""Prepare repository documentation/artifacts after real native completion.

No commit or upload; immutable released version artifacts remain untouched.
"""
from pathlib import Path
import json,shutil,hashlib
from prepare_inputs import ROOT,REPO,SOURCE,SOURCE_SHA

def edit(path,old,new):
    text=path.read_text(encoding='utf-8-sig')
    assert text.count(old)==1,(path,old)
    path.write_text(text.replace(old,new),encoding='utf-8',newline='\n')

def run():
    summary=json.loads((ROOT/'results/final_summary.json').read_text(encoding='utf-8'))
    assert summary['all25_native_executed'] and len(summary['results'])==25
    assert hashlib.sha256(SOURCE.read_bytes()).hexdigest()==SOURCE_SHA
    fs,targets=summary['fs_obtained'],summary['target_met']
    failed=[r['case'] for r in summary['results'] if not r.get('fs_obtained')]
    destination=REPO/'validation/G1R12_native25_20261009';assert not destination.exists()
    destination.mkdir(parents=True)
    for src,dst in [('FINAL_REPORT.md','REPORT.md'),('FINAL_RESULTS.csv','RESULTS.csv'),
                    ('source_identity.json','source_identity.json'),('environment.json','environment.json'),
                    ('Test_Case_Catalog.md','Test_Case_Catalog.md')]:
        shutil.copy2(ROOT/src,destination/dst)
    report=(destination/'REPORT.md').read_text(encoding='utf-8').replace('FINAL_RESULTS.csv','RESULTS.csv')
    (destination/'REPORT.md').write_text(report,encoding='utf-8',newline='\n')
    shutil.copytree(ROOT/'inputs',destination/'inputs')
    (destination/'results').mkdir()
    for name in ['final_summary.json','template_check.json','solve_costs.json','solve_costs.csv',
                 'native25_terminal.json','invariants.json','pure_c0_root_review.json','preservation_checks.json']:
        shutil.copy2(ROOT/'results'/name,destination/'results'/name)
    for name in ['surcharge_diagnosis','underflow_diagnosis','pure_c0_feasibility','pure_c0_upper_selection']:
        (destination/name).mkdir()
        for p in (ROOT/name).iterdir():
            if p.is_file() and p.suffix in ['.md','.json','.png','.svg','.py']:
                shutil.copy2(p,destination/name/p.name)
        if name=='pure_c0_upper_selection':
            (destination/name/'positive_control_Fs1_30').mkdir()
            shutil.copy2(ROOT/name/'positive_control_Fs1_30/result.json',destination/name/'positive_control_Fs1_30/result.json')
    scripts=destination/'scripts';scripts.mkdir()
    for p in ROOT.glob('*.py'):shutil.copy2(p,scripts/p.name)
    shutil.copytree(ROOT/'tests',scripts/'tests',ignore=shutil.ignore_patterns('__pycache__'))
    packages=json.loads((ROOT/'delivery/package_summary.json').read_text(encoding='utf-8'))
    for p in packages:
        source=ROOT/'delivery'/p['name'];assert hashlib.sha256(source.read_bytes()).hexdigest()==p['sha256']
        assert not (REPO/'downloads'/p['name']).exists()
        shutil.copy2(source,REPO/'downloads'/p['name'])
    shutil.copy2(ROOT/'delivery/SHA256SUMS.txt',destination/'SHA256SUMS.txt')
    sums=REPO/'downloads/SHA256SUMS.txt'
    sums.write_text(sums.read_text(encoding='utf-8').rstrip()+'\n'+(ROOT/'delivery/SHA256SUMS.txt').read_text(encoding='utf-8'),encoding='utf-8',newline='\n')
    proof=f'間隙水圧を除く25ケースすべてを実Excelで実行しました。監査済みFs区間は{fs}/25、GAP 1%と根探索目標の達成は{targets}/25です。最終Fs未取得は{"・".join(failed)}です。'
    short=proof+' [実行結果・制約](validation/G1R12_native25_20261009/REPORT.md)を確認してください。'
    readme=REPO/'README.md'
    edit(readme,'**G1R12の全P06N解析時間、GAP 1%到達、25ケース全回帰は未確認です。図の矢印は規準化された速度機構であり、変位量ではありません。**',
         '**'+short+' 図の矢印は規準化された速度機構であり、変位量ではありません。**')
    edit(readme,'[25ケース台帳（NOT_RUN）](versions/G1R12/excel_acceptance_g1r12.csv)',
         '[25ケースの実行結果](validation/G1R12_native25_20261009/RESULTS.csv) / [Issue修正時点の台帳（保全）](versions/G1R12/excel_acceptance_g1r12.csv)')
    links='\n'.join('- ['+label+'](https://github.com/tak063495-prog/excel-vba-rpfem/raw/refs/heads/main/downloads/'+p['name']+')' for label,p in zip(['25ケースの原ログ・監査・再現コードZIP','25ケースの結果XLSM ZIP'],packages))
    edit(readme,'**全P06N・GAP1%・25ケース全回帰・全支持力/剛体回帰は未実施です。** G1R11の公開成果物は保全しています。',
         '### 2026年10月9日の25ケース実行\n\n'+proof+' '+f'取得Fsのうち{fs-targets}ケースは精度目標が未達です。全支持力・剛体回帰は今回の対象外です。\n\n'+links+'\n- [報告書](validation/G1R12_native25_20261009/REPORT.md) / [結果CSV](validation/G1R12_native25_20261009/RESULTS.csv)\n\n製品VBAは同じG1R12で、試験用証拠モジュールを追加したコピーを使いました。D07Aの荷重端点メッシュ、D03NのDD積underflowについてPythonで原因・対策を検証しましたが、修復のVBA実装はまだ行っていません。G1R11/G1R12の既存公開XLSM・ZIPは保全しています。')
    usage=REPO/'docs/USAGE_JA.md'
    edit(usage,'全P06N・25ケース全回帰は未実施です。',proof+' [実行結果](../validation/G1R12_native25_20261009/REPORT.md)を参照してください。')
    notes=REPO/'docs/NOTES_JA.md'
    edit(notes,'G1R12の全P06N、GAP1%、25ケース、全支持力・剛体回帰はNOT_RUNです。下記の1000要素固定点などはG1R11時点の証拠で、G1R12で全てを再実行したという意味ではありません。[G1R12試験結果](../versions/G1R12/TEST_RESULTS.csv)と[25ケース台帳](../versions/G1R12/excel_acceptance_g1r12.csv)を参照してください。',
         proof+' 取得Fsを全て精度目標達成と扱いません。[25ケース結果](../validation/G1R12_native25_20261009/REPORT.md)を参照してください。全支持力・剛体回帰は今回の対象外です。下記の固定点はG1R11の保全した証拠です。G1R12の[Issue修正時点の試験結果](../versions/G1R12/TEST_RESULTS.csv)と[当時の台帳](../versions/G1R12/excel_acceptance_g1r12.csv)も原形で残しています。')
    attrs=REPO/'.gitattributes'
    attrs.write_text(attrs.read_text(encoding='utf-8').rstrip()+'\nvalidation/G1R12_native25_20261009/** -text\n',encoding='utf-8',newline='\n')
    assert hashlib.sha256(SOURCE.read_bytes()).hexdigest()==SOURCE_SHA
    print(json.dumps({'prepared':True,'fs_obtained':fs,'target_met':targets,'failed':failed,'path':str(destination)},ensure_ascii=False))

if __name__=='__main__':run()
