"""Japanese release report, generated only after executed evidence passes."""
import csv,json
from prepare import ROOT,write,sha

def read(name):return json.loads((ROOT/'results'/name).read_text(encoding='utf8'))
def main():
 c=read('comparison.json');v=read('delivery_verification.json');fixed=read('fixed_comparison.json');roots=read('root_regressions.json')
 core=read('core_pipeline.json');pipeline=read('pipeline.json');dll=sha(ROOT/'delivery/rpfem_native_x64.dll')
 assert core['state']==pipeline['state']=='completed' and len(core['runs'])==22 and len(pipeline['runs'])==8
 assert v['reopen_whole_project_compile']=='PASS' and v['protected_original_files']==1085
 assert all(r['exit_code']==0 for r in core['runs']+pipeline['runs'])
 assert fixed['passed'] and roots['passed_comparisons']==4
 checks=['bulk_oracle/summary.json','lu_oracle/summary.json','nt_oracle_extended/comparison.json','boundary_oracle/summary.json','job_contract/summary.json','new_job_contract/summary.json','job_lifecycle.json','new_job_lifecycle.json','schur_python.json','krylov_python.json','hdd_python.json','compile_python.json','guard_oracle/summary.json','mode_switch_final/summary.json']
 for name in checks:assert read(name)['dll_sha256']==dll,name
 assert all(dll==r['DLL5_sha256'] for r in roots['cases']) and c['dll_sha256']==dll==fixed['final_DLL5_sha256']
 dd=read('dd_core_python.json');assert dd['DLL5_sha256']==dll and dd['bitwise_exact_successes_and_identical_error_precedence']
 nt=read('nt_oracle_extended/comparison.json');assert nt['bitwise_exact_rows']==174 and not nt['classification_mismatches'] and not nt['differences']
 lu=read('lu_oracle/summary.json');assert lu['exact_factor_and_solution']==352 and lu['singular_identical']==8 and not lu['failures']
 assert read('bulk_oracle/summary.json')['exact'] and read('boundary_oracle/summary.json')['bitwise_exact']
 assert read('job_contract/summary.json')['passed']==12 and read('new_job_contract/summary.json')['passed']==24
 assert read('issue_regression/summary.json')['passed']==51 and read('issue_regression/summary.json')['failed']==0
 assert read('fallback_seep.json')['same_hsd_result_all_backends']
 assert all(row['exit_code']==0 for row in read('native_faults.json').values())
 times=c['seconds'];base=times['VBA_existing']
 timing_rows='\n'.join(f"| {label} | {times[key]:.3f} | {base/times[key]:.2f}倍 | {100*(1-times[key]/base):.2f}% | {kind} |" for key,label,kind in [('VBA_existing','DLLなし','既存実測ログ'),('DLL3_existing','DLL3','既存実測ログ'),('DLL4_fresh','DLL4','今回の再実測'),('DLL5_fresh','DLL5','今回の実測')])
 fixed_rows='\n'.join(f"| {key} | {r['elements']} | {r['Fs_input']:g} | {r['seconds']['vba']:.3f} | {r['seconds']['dll4']:.3f} | {r['seconds']['dll5']:.3f} | {r['reduction_percent_vs_DLL4']:.2f}% |" for key,r in fixed['FEM'].items())
 root_rows='\n'.join(f"| {r['case']} | {r['terminal_status']} | {format(r['lower'],'.15g') if r['Fs_obtained'] else '未取得'} | {format(r['upper'],'.17g') if r['Fs_obtained'] else '未取得'} | {r['seconds']['DLL4_existing']:.2f} | {r['seconds']['DLL5_current']:.2f} | {r['reduction_percent_vs_DLL4']:.2f}% |" for r in roots['cases'])
 text=f'''# G1R16 DLL5 検証報告

実Excel 16.0・64bit、同一PCで検証しました。DLL4を基準に、Schur内部処理、HSDのDD組立とGMRES修復、疎パターン準備、入力比較、時間区間ごとの一括演算、DD/疎Double演算の呼出し削減をDLL5へ実装しました。予定した主要な追加最適化をDLL5にまとめています。物理モデル・Fs定義・収束条件・救済方針は維持しています。

## 堤防例題の速度

**DLL4 {times['DLL4_fresh']:.3f}秒 → DLL5 {times['DLL5_fresh']:.3f}秒、{c['reduction_percent_vs_DLL4']:.2f}%短縮、約{c['speedup_vs_DLL4']:.2f}倍**でした。

| 版 | 解析時間 秒 | DLLなし基準の倍率 | DLLなし基準の短縮率 | 時間の由来 |
|---|---:|---:|---:|---|
{timing_rows}

DLLなし・DLL3は同一PCの既存実測記録です。ユーザーの指示に従って長時間のVBA速度測定は再実行していません。DLL4とDLL5はそれぞれ新しいExcelプロセスで順番に実行しました。この時間測定と本作業の別の数値計算は重ねていません。各版1回の全解析比較で、すべてのPCやケースに対する倍率の保証ではありません。

同一入力はTARGET=600、q=8、適応ON、kh=0.15右、浸透連成ONです。表は`RPX_Run`のログの`run_end`までで、別途実行した浸透・メッシュ作成、保存、Python照合は含みません。COM呼出し全体はDLL4 {c['DLL4_COM_seconds']:.6f}秒、DLL5 {c['DLL5_COM_seconds']:.6f}秒でした。線形核だけの改善率を全解析の倍率として扱っていません。

Fs下界 **{c['lower']:.17g}**、上界 **{c['upper']:.17g}** は保存済みDoubleとしてDLL4および既存VBA/DLL3とビット一致しました。入力・材料・設定・浸透入力・節点・要素の保存値も一致しました。53求解の探索順・Fs点・反復数・目的値・残差・GAP・判定、603監査の詳細はDLL4および既存VBA/DLL3とログ記録精度で一致しています。ログ値の一致と全解ベクトルのビット一致は区別し、全ベクトルは次の固定点試験で検証しました。

証跡: `results/comparison.json`、`results/levee_dll4_serial/`、`results/levee_dll5_final/`。旧VBA/DLL3の参照元とSHAも記録しています。

## 固定Fsの実FEM比較

TARGET=256・q=8・適応OFF・浸透OFF・kh=0で、VBAとDLL4の実行結果を保存し、最終DLL5の同じ6点を照合しました。**解の全成分・DD出力・メッシュのバイナリがVBA/DLL4/DLL5で完全一致し、独立Python物理監査も合格**しました。1バックエンドだけの自己比較は精度の証拠として扱っていません。

| 固定点 | 実要素数 | 入力Fs | VBA 秒 | DLL4 秒 | DLL5 秒 | DLL4比短縮 |
|---|---:|---:|---:|---:|---:|---:|
{fixed_rows}

表のFsは求まった最終安全率ではなく固定入力です。C5_45はc=5・φ=ψ=45、P06A/P06Nは多層の関連/Davisモデルです。P06下界の固定Fs=1で荷重倍率が非常に小さくなる現象も基準版と同じです。これを最終Fsが得られたという意味にしません。

証跡: `results/fixed_comparison.json`、`results/compiled_q8_first/`、`results/fixed_q8_release/`。

## 回帰・ロバスト性試験

| 試験 | 実行形態・結果 | 証跡 |
|---|---|---|
| 主回帰 | 22試験グループを逐次実行、すべて終了コード0 | `results/core_pipeline.json` |
| 入力比較 | 実Excel18条件。型、Empty、CVErr、文字列、数式、events OFF時の変更 | `guard_oracle/summary.json` |
| HSD状態の同一性 | 実Excel12条件。1ULP、tau/kappa、符号付き0、NaN payload、配列境界 | 同上 |
| HSD既知解 | 実VBA/DLL5の6組と実DLL3参照の出力が完全一致 | `fixed_comparison.json`、`hsd_regression/` |
| LU | 実VBA/DLL5、360条件。因子・ピボット・解352組一致、特異8は同じ拒否。独立Python最大後退誤差{lu['max_independent_backward_error']:.3e} | `lu_oracle/summary.json` |
| NT | 実VBAとネイティブAPI、極端な値を含む224入力。成功174はhi/loも一致、拒否50の判定一致 | `nt_oracle_extended/comparison.json` |
| 一括演算 | 実VBA/DLL5、160組・1,239,592バイト一致 | `bulk_oracle/summary.json` |
| LDL/DD/Schur・順列 | 実VBA/DLL5。特異、短い配列、overflow、取消を確認。順列68組一致、n>20000を含む | 同上 |
| コーン境界 | 実VBA/DLL5、5070条件。異常時の途中出力と例外も一致 | `boundary_oracle/summary.json` |
| 既存不具合 | 実Excel、51アサーション合格 | `issue_regression/summary.json` |
| 既存ジョブの障害 | 実Excel、取消/OOM/添字/未知例外/入力変更/成功の12条件 | `job_contract/summary.json` |
| 新規ジョブの障害 | 実Excel、compiler/Schur factor/Schur solve/Krylovの24条件。例外18/7/9/513/5と解放を確認 | `new_job_contract/summary.json` |
| C++例外 | 試験実行ファイルを実ビルド・実行。新compiler211・Schur30・Krylov17の確保失敗と各4未知例外、既存核も検査 | `native_faults.json` |
| Schur独立検証 | Python行列・独立残差240ジョブ、最大後退誤差{read('schur_python.json')['maximum_independent_backward_error']:.3e} | `schur_python.json` |
| Krylov独立検証 | Python/NumPy独立解80ジョブ、最大後退誤差{read('krylov_python.json')['max_independent_backward_error']:.3e} | `krylov_python.json` |
| HSD DD組立 | Python 80桁Decimal、60条件 | `hdd_python.json` |
| 疎構造 | Pythonで120条件。署名全文、CSC、slot、挿入順、既存並べ替えと比較 | `compile_python.json` |
| DD融合演算 | DLL4/DLL5の7552条件。成功ビットと異常判定・優先順一致 | `dd_core_python.json` |
| 大きいジョブの継続 | ネイティブAPIで実際の時間分割・再開・再入/古い世代拒否・途中解放を検査。新Schur98305変数、Krylov65536変数 | `new_job_lifecycle.json`、`job_lifecycle.json` |
| モード切替 | 実Excel、auto→VBA/autoの世代無効化と新規求解 | `mode_switch_final/summary.json` |
| DLL不在・旧DLL | 実Excel、missing/旧DLL2/旧DLL3/旧DLL4のVBA復帰と浸透/水位試験 | `fallback_seep.json` |
| 最終XLSM | 実Excelで保存、再読込、70モジュール全コンパイル、DLL_READY、4ボタンを確認 | `delivery_verification.json` |

Pythonの独立残差検証、ネイティブAPIの契約試験、実Excelの比較・障害注入、全解析を区別しています。故障注入コードは試験コピーのみに配置しました。配布ブックに試験モジュールを含めていません。

## 難しい4ケースの粗メッシュ比較

TARGET=200・q=8・適応OFF、元カタログの形状/多層材料、重力固定、浸透/地震OFFです。DLL4の保存済み粗メッシュ結果と同じ終端まで照合しました。

| ケース | DLL5終端 | Fs下界 | Fs上界 | DLL4 秒 | DLL5 秒 | DLL4比短縮 |
|---|---|---:|---:|---:|---:|---:|
{root_rows}

D03N/D07Aの`PARTIAL_CERTIFIED`はFs区間が得られても目標精度は未達という判定です。P06Aはユーザーの選択どおりFs≥0.125の最初の15完了求解だけを比較し、試験コピーのみで次の点の前に停止しました。P06Nの粗メッシュ下界は基準DLL4と同じ`FS_LOWER_BRACKET_NOT_FOUND`でFs未取得です。**この回帰一致をケース合格やロバスト性問題の解決として扱っていません。**

4組とも探索点・求解/監査記録・入力/メッシュが一致し、Fs取得2組は保存値もビット一致しました。P06Aの時間は15点までのログ、その他は粗メッシュRPX_Runの呼出し時間です。全適応解析の速度を推定する表ではありません。証跡: `results/root_regressions.json`。

P06A/P06Nのフル適応解析、P06Aの極小Fs探索、以前中断した26ケース全試験は再開していません。今回の速度からP06Aの約10時間の完走時間を保証することはできません。25/26ケースすべての安全率が出ることも今回の合格条件にしていません。

## 試験停止と開発中の観測

回帰の最初のNT比較は、Excelのx87設定`0x137f`に対してPythonが`0x027f`であり、不一致で停止しました。Excelの設定を読み取り、比較用Pythonプロセスだけを合わせて再実行すると188条件/追加224条件の成功ビットと拒否判定が完全一致しました。製品DLLによる浮動小数点環境の変更や許容誤差の緩和は行っていません。単独起動でも比較前に設定を観測するよう試験を修正しました。

追加の障害試験の再実行では前回出力の存在を正常に拒否した`OUTPUT_EXISTS`で停止しました。出力先を実行ごとに分け、24条件を最終DLLで再実行しました。合格済み試験の証跡を残して再開しています。停止・再開と元の失敗ログは`core_pipeline.json`、`core_*_previous_*.log`に残しています。

開発中、Excel導入/試験プロセスの後処理が重なった時点でExcel終了時のntdll例外を1回観測しました。因果は未確定です。以後はExcelの開始から終了まで逐次実行し、最終の回帰・計算・再読込で同じクラッシュは再現していません。

## 保全・配布

基準DLL4 commitは`6f9c2248945f2cda25fd3a04b40ab962a70187bf`です。元の1085ファイルをSHA-256で保全し、既存8標準モジュールを変更、新規1モジュールを追加しました。既存61モジュールとdocument moduleは無変更です。インポート用9ファイルはCP932/CRLFを厳密検証しています。旧P2-Sの32項目も別途検査します。

最終DLL SHA-256: `{dll}`。最終XLSM SHA-256: `{v['xlsm_sha256']}`。

実Excelで保存した結果ブックをバイト単位でコピーし、再読込・全コンパイルを確認しました。原本変更、vbaProject.bin差し替え、マクロブックのライター書戻し、セキュリティ変更は行っていません。導入・復旧は`IMPORT_ROLLBACK_JA.md`、コード対応は`IMPLEMENTATION_JA.md`、試験の一覧は`TEST_RESULTS.csv`、再現手順は`tools/REPRODUCTION_JA.md`を参照してください。
'''
 (ROOT/'REPORT.md').write_text(text,encoding='utf8',newline='\n')
 readme=f'''# G1R16 DLL5

通常内点法とHSDを含む追加最適化版です。堤防例題はDLL4 {times['DLL4_fresh']:.2f}秒からDLL5 {times['DLL5_fresh']:.2f}秒へ、{c['reduction_percent_vs_DLL4']:.2f}%短縮しました。Fs・入力・メッシュは一致し、全53求解/603監査の記録も一致しました。同一PC各1回の実測です。

`delivery/`のXLSMとDLLを組で同じフォルダーに置きます。DLL交換前にExcelを終了し、原本を残して別名で保存してください。DLLがない場合・非対応版の場合は従来VBAで計算します。マクロの利用は組織の規則に従ってください。通常の解析にPythonやVBAプロジェクトアクセスの許可は不要です。

配布XLSMはTARGET=600、q=8、適応ON、kh=0.15右、浸透ONの堤防例題です。「要素定義」の「解析」で再計算できます。Fs下界{c['lower']:.15g}、上界{c['upper']:.17g}の結果を保存しています。

- [検証報告・速度比較](REPORT.md)
- [実装範囲・コードの対応](IMPLEMENTATION_JA.md)
- [導入・復旧](IMPORT_ROLLBACK_JA.md)
- [試験結果](TEST_RESULTS.csv) / [再現手順](tools/REPRODUCTION_JA.md)
- [VBA差分](changes.diff) / [ネイティブ差分](native_changes.diff)
- [編集用ソース](src/vba/) / [VBEインポート用](import/) / [DLLソース・ビルド](native/)

P06A/P06Nは固定Fs比較と限定した粗メッシュ比較です。フル適応解析、極小Fs探索、26ケース全再試験は未実施です。粗メッシュP06Nの基準版と同じFs未取得も記録しています。`PARTIAL_CERTIFIED`と目標精度達成を区別してください。

ABI `0x00160005`、Windows x64。ライセンスは`native/licenses/`、ファイルのハッシュは`SHA256SUMS.txt`を確認してください。
'''
 (ROOT/'README.md').write_text(readme,encoding='utf8',newline='\n')
 rows=[('core','mixed actual Excel/native API/Python','PASS','22 test groups','results/core_pipeline.json')]
 for label,kind,scope,path in [('Guard','actual Excel','18 input +12 exact HSD identity','guard_oracle/summary.json'),('HSD','actual VBA/DLL5 + recorded DLL3','6 known solutions','fixed_comparison.json'),('Fixed_FEM','actual VBA/DLL4/DLL5 + independent Python audit','6 q8 points, full fields and mesh bitwise exact','fixed_comparison.json'),('LU','actual VBA/DLL5 + independent Python residual','360','lu_oracle/summary.json'),('NT','actual VBA/native API','224:174 exact+50 matching rejects','nt_oracle_extended/comparison.json'),('Bulk','actual VBA/DLL5','160;1239592 bytes','bulk_oracle/summary.json'),('Boundary','actual VBA/DLL5','5070','boundary_oracle/summary.json'),('Existing_job_faults','actual Excel failure injection','12','job_contract/summary.json'),('New_job_faults','actual Excel failure injection','24','new_job_contract/summary.json'),('Native_faults','C++ actual build/execute','all allocation/unknown injection contracts','native_faults.json'),('Schur','independent Python','240','schur_python.json'),('Krylov','independent Python','80','krylov_python.json'),('DD_assembly','80-digit Decimal Python','60','hdd_python.json'),('Compiler','independent Python sparse construction','120','compile_python.json'),('DD_fused','DLL4/DLL5 native API','7552','dd_core_python.json'),('Large_lifecycle','native API','real yields/ownership/stale/cancel cleanup','new_job_lifecycle.json'),('Mode_switch','actual Excel','auto/vba and auto/auto','mode_switch_final/summary.json'),('Fallback_seep','actual Excel','native,missing,old2,old3,old4 +seep/water','fallback_seep.json'),('Issue1_2','actual Excel','51 assertions','issue_regression/summary.json'),('Levee','actual DLL4/DLL5 full RPX_Run; prior VBA/DLL3','53 solves603 audits; saved Fs bits exact','comparison.json'),('Delivery','actual Excel compile/save/reopen','70 modules,4 buttons,DLL_READY','delivery_verification.json')]:rows.append((label,kind,'PASS',scope,'results/'+path))
 for r in roots['cases']:
  rows.append((r['case']+'_regression','actual coarse DLL5 vs recorded DLL4','PASS',r['scope'],'results/root_regressions.json'))
  rows.append((r['case']+'_final_Fs','actual coarse DLL5',r['terminal_status'], 'Fs obtained' if r['Fs_obtained'] else 'Fs unavailable; do not count as pass','results/root_regressions.json'))
 rows += [('P06A_full','full adaptive Excel','NOT_RUN','user stopped full analysis','results/root_regressions.json'),('P06N_full','full adaptive Excel','NOT_RUN','not restarted','results/root_regressions.json'),('P06A_tiny_Fs','extreme Fs Excel','NOT_RUN','user limited floor0.125','results/root_regressions.json'),('26_case_batch','actual full cases','NOT_RUN','user interrupted batch earlier','results/root_regressions.json')]
 with (ROOT/'TEST_RESULTS.csv').open('w',encoding='utf-8-sig',newline='') as f:
  w=csv.writer(f);w.writerow(['test','validation_type','result','executed_scope','evidence']);w.writerows(rows)
 write(ROOT/'WORK_STATUS.json',{'state':'VALIDATED_READY_FOR_PUBLICATION','protected_original_files':1085,'final_dll_sha256':dll,'final_xlsm_sha256':v['xlsm_sha256'],'core_groups_completed':22,'final_groups_completed':8,'full_levee':'PASS','fixed_FEM_and_HSD':'PASS','coarse_comparisons':4,'not_run':['full P06A','full P06N','tiny-Fs exploration','interrupted 26-case batch'],'user_stop_preferences_preserved':True})
 print('REPORT evidence assertions PASS',flush=True)
if __name__=='__main__':main()
