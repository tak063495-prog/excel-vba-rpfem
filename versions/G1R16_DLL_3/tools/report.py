import json
from prepare import ROOT,write,sha
def read(name):return json.loads((ROOT/'results'/name).read_text(encoding='utf-8'))
def main():
    c=read('comparison.json');v=read('delivery_verification.json');bulk=read('bulk_oracle/summary.json');lu=read('lu_oracle/summary.json');nt=read('nt_oracle_extended/comparison.json');hsd=read('hsd_regression/summary.json');fem=read('fixed_fem_final/summary.json');issue=read('issue_regression/summary.json');fallback=read('fallback_seep.json')
    dllhash=sha(ROOT/'delivery/rpfem_native_x64.dll')
    assert all(x['dll_sha256']==dllhash for x in [c,v,bulk,lu,nt,fem])
    assert bulk['exact'] and not lu['failures'] and not nt['differences'] and not nt['classification_mismatches']
    assert len(hsd['cases'])==6 and all(x['bitwise_exact'] for x in hsd['cases'])
    assert len(fem['cases'])==6 and all(x['same_mesh'] and x['bitwise_same_field'] and x['vba']['independent_audit']['pass'] and x['auto']['independent_audit']['pass'] for x in fem['cases'])
    assert issue['failed']==0 and issue['passed']==51 and fallback['same_hsd_result_all_backends']
    phases=[('nt_scaling','NTスケーリング'),('kkt_update','KKT組立'),('corrector','Predictor–corrector補正'),('gproduct','G積'),('gtadd','G転置積'),('linear_solve_and_refinement','線形求解・補正'),('ordering','並べ替え')]
    table='\n'.join(f"| {label} | {c['previous_phases'][key]:.3f} | {c['native_phases'][key]:.3f} |" for key,label in phases)
    femtable='\n'.join(f"| {x['case']} | {x['side']} | {x['Fs']} | {'97' if x['case'].startswith('P06') else '71'} | ビット一致 | 合格 |" for x in fem['cases'])
    report=f'''# G1R16 DLL_3 検証報告

## 実測と精度

2026年10月12日、実Excel 16.0・64bitで検証。堤防例題の解析時間は **DLL2 {c['previous_DLL2_seconds']:.2f}秒 → DLL_3 {c['native_seconds']:.2f}秒**、**{c['reduction_percent_vs_DLL2']:.2f}%短縮**でした。
同じPCの既存VBA記録 {c['VBA_existing_seconds']:.2f}秒に対して **{c['reduction_percent_vs_existing_VBA']:.2f}%短縮・約{c['VBA_existing_seconds']/c['native_seconds']:.2f}倍**です。ユーザー指示に従い、VBA版の長時間測定は再実行していません。

| 指標 | G1R16 VBA（既存記録） | G1R16 DLL2（既存記録） | DLL_3（今回） |
|---|---:|---:|---:|
| 解析時間（秒） | {c['VBA_existing_seconds']:.2f} | {c['previous_DLL2_seconds']:.2f} | {c['native_seconds']:.2f} |
| Fs下界 | 1.2244477290215 | 1.2244477290215 | {c['lower']:.16g} |
| Fs上界 | 1.23430802878319 | 1.23430802878319 | {c['upper']:.17g} |
| 求解回数 | 53 | 53 | 53 |
| 最終メッシュ | 596要素 | 596要素 | 596要素 |
| 監査ログ | 603件 | 603件 | 603件 |

条件はTARGET=600、最適化ON、kh=0.15右、浸透連成ON。時間は`RPX_Run`の`run_end`までの記録です。事前浸透、事前メッシュ、ファイル保存は含みません。COMから測った時間は{c['com_wall_seconds']:.3f}秒です。単一PCの既存記録との比較であり、交互反復測定ではありません。

実G1R16の保存済みDLL2結果とのFsはビット一致。全53求解について、探索点・求解順・反復数・目的値・残差・GAPがVBA記録およびDLL2記録とログ記録精度で一致しました。603件の物理監査も全詳細が一致しました。入力4シート、節点・要素の保存済みデータも一致しています。

G1R16の配布元XLSMにはG1R15時点の計算結果・メッシュが残っていました。また、G1R16の`REPORT.md`には上界1.2344080288という転記があります。実G1R16のログ・結果ブックは1.2343080287831916です。比較基準には実行済みG1R16のデータを使用し、元の配布XLSM・報告書は変更していません。参照証跡は`results/reference/`に収録しました。

## 相別の時間

| 処理 | DLL2（秒） | DLL_3（秒） |
|---|---:|---:|
{table}

相別記録には重複する測定範囲があるため、合計は全解析時間と一致しません。今回の追加高速化はNT・KKT・コーン積・HSD原作用素などの配列内ループをC/ASMへ移した効果です。並べ替えやLDLなど、DLL2で既に移植した処理は継承しています。

## 実VBAとの数値比較

| 試験 | 結果 | 証跡 |
|---|---|---|
| NT単体・極端なスケール・境界点 | 224入力。成功174入力の全DoubleとDD hi/loがビット一致。拒否50入力の例外番号も一致 | `results/nt_oracle_extended/` |
| 一括NT、補正、G/G転置、KKT、HSD DD原作用素 | 160組、{bulk['bytes_vba']:,}バイトの出力が完全一致。シフト有無、精密NT選択、係数の更新を含む | `results/bulk_oracle/` |
| HSD LU | 360条件。352条件のピボット・因子・解が完全一致。特異8条件も同じ判定 | `results/lu_oracle/` |
| LUのPython独立残差 | 最大後退誤差 {lu['max_independent_backward_error']:.3e} | 同上 |
| HSD求解 | LDL_CHECKEDとPIVOT_ONLYの6組で、解・反復数・残差がビット一致。既知の厳密解とも一致 | `results/hsd_regression/` |
| DLL2から継承した数値核 | Double/DD/Schur、特異ピボット、短い配列、オーバーフロー、取消が合格 | `results/bulk_oracle/summary.json` |
| 並べ替え | 68組の順列・逆順列が完全一致。n>20000、同次数、挿入順、取消、不正IDを含む | 同上 |
| 原本・保存済みソース | 元版198ファイルのSHA-256不変、68モジュール確認、59既存モジュール不変、文書モジュール不変 | `results/delivery_verification.json` |

これらは実Excelで元のVBAを実行した比較です。Pythonだけの模擬計算をVBA実行済みとは扱っていません。Pythonは独立残差・物理監査と、実VBA出力とのバイナリ比較に使用しました。

## 固定Fsでの実FEM比較

| 入力 | 上下界 | 設定Fs | 要素数 | VBA対DLL_3の計算場 | Python物理監査 |
|---|---|---:|---:|---|---|
{femtable}

C5_45はc=5・φ=ψ=45の制御試験です。P06A/P06Nは多層・薄いc=0弱層を含むモデルです。TARGET=64・q=2・間隙水圧なしで、同じメッシュと材料・荷重を使用しました。P06の4求解では実際にHSD_C0経路、下界ではDD-LDLも使用されました。

**設定Fs=1は試験の入力であり、P06の最終安全率ではありません。** この固定点での下界荷重倍率はP06Aが約7.49E-11、P06Nが約8.56E-12で、元VBAと一致しています。今回P06のフルメッシュ・適応・最終Fs探索を再実行したとは扱いません。固定点試験の壁時計は準備・順序再利用・並行検証の影響を含むため、全ケースの速度倍率へ換算しません。

## ロバスト性・既存不具合の回帰

- HSD LUの省略引数、因子共有、epoch、メモリ上限、ピボットしきい値、有限性判定、取消が合格。
- C++の異常注入：LUで66割当点、並べ替えで2835割当点のメモリ不足を試験。各8点の未知例外も試験し、例外コードの保持とコンテキスト解放を確認。
- Issue #1/#2の51項目が合格。異なるメッシュ・場の所有者を混在させず、取消・メモリ不足・添字異常・未知例外・監査違反を通常のHSD救済に回さないことを確認。
- DLL不在と旧DLLでは自動的にVBAへ切替。HSDのLU経路を含む結果が、DLL_3使用時とビット一致。
- G1R16の浸透回帰：試験ダムQ=2.4181、緩和1E-8の未収束、緩和0.5/1.0の収束、地下水面・逆算自己試験が合格。

VBAの有限性判定と例外を保持し、NTのべき乗をVBAの実行時ライブラリへ合わせました。強度更新時のコーン配列の無効化も上下界の両方に実装しています。入力表の確認は既存設定の間隔で行い、取消と`DoEvents`は処理の区切りごとに維持しました。

## 成果物と範囲

`delivery/`に実解析結果を保存したXLSMとDLL、`src/vba/`に編集用UTF-8ソース、`import/`にCP932・CRLFの取込ファイル、`changes.diff`に差分、`tests/`と`tools/`に検証・再現用コード、`results/`に証跡があります。

32bit Excelでの実行と、P06のフルFs探索・26ケース全試験は今回未実施です。以前に中断された26ケースの一斉試験は再開していません。今回の合格範囲は上記の実行済み試験です。
'''
    (ROOT/'REPORT.md').write_text(report,encoding='utf-8')
    write(ROOT/'results/final_tests_index.json',{'DLL_3':dllhash,'native_excel_compile':'PASS','NT_rows':nt['rows'],'NT_success_bitwise':nt['bitwise_exact_rows'],'bulk_cases_bitwise':160,'LU_cases':360,'LU_bitwise':352,'LU_singular_same':8,'HSD_solves_bitwise':6,'fixed_FEM_pairs_with_python_audit':6,'ordering_pairs_bitwise':68,'issue_assertions_PASS':51,'missing_and_old_DLL_HSD_fallback':'PASS','seep_regression':'PASS','full_levee_solves':53,'full_levee_audits_exact':603,'protected_original_files':198,'32bit_Excel':'NOT_RUN','full_P06_Fs_search':'NOT_RUN','prior_interrupted_native26_batch':'NOT_RESUMED'})
    print('REPORT and test index written',flush=True)
if __name__=='__main__':main()
