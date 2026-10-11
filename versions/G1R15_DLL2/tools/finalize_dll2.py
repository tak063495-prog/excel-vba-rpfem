from pathlib import Path
import json,hashlib,shutil,subprocess,sys,zipfile
ROOT=Path(__file__).resolve().parents[1]
run=ROOT/'results/dll_run1_ordering'
status=json.loads((run/'status.json').read_text(encoding='utf-8'))
assert status['status']=='completed'
dst=ROOT/'delivery/RPFEM_20261012_G1R15_DLL2_Levee_Checked.xlsm'
shutil.copyfile(run/'RPFEM_G1R15_DLL_Levee.xlsm',dst)
assert hashlib.sha256((ROOT/'delivery/rpfem_native_x64.dll').read_bytes()).hexdigest()==status['dll_sha256']
for script in ['compare_results.py','verify_delivery.py','verify_previous.py','check_missing_order.py','check_saved_delivery.py']:
    subprocess.run([sys.executable,'-X','utf8',str(ROOT/'tools'/script)],check=True,cwd=ROOT.parent.parent)
comparison=json.loads((ROOT/'results/comparison.json').read_text(encoding='utf-8'))
final=next(r for r in comparison['runs'] if r['tag']=='dll_run1_ordering')
assert final['audit_details_exact_at_logged_precision']
assert final['solve_end_status_iteration_objective_residual_gap_exact_at_logged_precision']
assert final['previous_dll_audits_exact_at_logged_precision'] and final['previous_dll_solve_ends_exact_at_logged_precision']
assert all(final['previous_dll_saved_inputs_and_mesh_exact'].values())
assert all(not final['saved_book_comparisons'][name] for name in ['要素定義','材料データ','設定','浸透入力'])
assert final['saved_book_comparisons']['節点データ'] and final['saved_book_comparisons']['要素データ']
tests=json.loads((ROOT/'results/native_kernel_tests.json').read_text(encoding='utf-8'))
assert all(v.startswith(('PASS','DLL_READY')) for v in tests.values())
verification=json.loads((ROOT/'results/delivery_verification.json').read_text(encoding='utf-8'))
phase_names=[('ordering','並べ替え'),('convert_pattern','疎行列への変換'),('numeric_ldl','数値LDL'),('linear_solve_and_refinement','線形求解・補正'),('nt_scaling','NTスケーリング')]
phase_table='\n'.join(f"| {label} | {final['previous_dll_phases'][name]:.2f}秒 | {final['dll_phases'][name]:.2f}秒 |" for name,label in phase_names)
report=f'''# G1R15 DLL2：並べ替え移植による追加高速化

## 実測結果

同じ堤防例題の解析時間は、前回DLL版 **{comparison['previous_dll_seconds']:.2f}秒 → {final['analysis_seconds']:.2f}秒**。
前回から **{final['reduction_percent_vs_previous_dll']:.2f}%短縮、{final['speedup_vs_previous_dll']:.3f}倍**。
元VBAの既存記録 **{comparison['reference_seconds']:.2f}秒**との比較では **{final['reduction_percent']:.2f}%短縮、{final['speedup']:.3f}倍**。
VBA版はユーザー指示により再計測していない。前回DLL版・元VBA版はこのPCの既存記録を使用した。
交互測定ではなく、他のPC・ケースでの倍率は保証しない。

| 指標 | 元VBA（既存記録） | 前回DLL（既存記録） | DLL2（今回実行） |
|---|---:|---:|---:|
| run_end時間 | {comparison['reference_seconds']:.6f}秒 | {comparison['previous_dll_seconds']:.6f}秒 | {final['analysis_seconds']:.6f}秒 |
| Fs下界 | 1.22444773581456 | 1.22444773581456 | {final['lower']:.15g} |
| Fs上界 | 1.23430803570109 | 1.23430803570109 | {final['upper']:.15g} |
| 上下界差 | 約0.802056% | 約0.802056% | {final['gap_percent']:.6f}% |
| 求解回数 | 53 | 53 | {final['solve_count']} |
| 終端 | OPTIMAL | OPTIMAL | OPTIMAL |

時間は `RPX_Run` 内の記録値。浸透の事前計算、最初のメッシュ作成、保存は含まない。
それらは `results/dll_run1_ordering/status.json` に別記録した。

## 追加した処理

前回DLL版で約90秒かかっていた `RPX_Order` の商グラフ最小次数並べ替えを、C++へ移植した。
一般LU用の `RPX_ColumnOrder` も同じ処理基盤で実装した。
並べ替え手法、キーの挿入順、削除と再登録の順序、同次数でのヒープ選択、n>20000分岐を維持した。
整数処理だけを追加移植し、浮動小数点の計算法・収束基準は前回DLL版のままとした。

| 変更対象 | 内容 |
|---|---|
| RPX_Order | DLLが利用できる場合の入口を追加。既存VBA本体を保持 |
| RPX_ColumnOrder | 暗黙の正規グラフを移植。AᵀA行列を明示生成しない。既存VBA本体を保持 |
| RPX_Native | グラフ転送・分割処理・解放・ABI確認・例外伝達を追加 |
| rpfem_order.cpp | 挿入順集合、ヒープ、商グラフのC++実装 |
| rpfem_native.c | 前回の数値核を保持。対応ABIを0x00150002へ更新 |
| RPX_LDL / RPX_SchurLDL / RPX_SeepUI | 前回DLL版と同じ。元G1R15との差分も同梱 |

MC、Davis変換、Fs定義、上下界定式化、荷重、NT/HSD反復、許容差、探索、適応政策、物理監査を変更していない。
入力・メッシュ・荷重のキャッシュ判定も変更していない。

## 精度とロバスト性の確認

- **68組**の実Excel比較で、VBA/DLLの並び順・逆順表が整数単位で完全一致。
  対称グラフと暗黙の正規グラフ、1/2/7/31/70/97/270/20481行、空・疎・密・非対称・同次数・再登録を含む。
- 堤防の**53求解**すべての候補・Fs試行値・メッシュ条件が、元VBAログと一致。
- **物理監査603記録すべてPASS**。詳細が元VBA・前回DLLのログと一致。
- 全求解の終端状態・反復番号・目的値・主双対残差・ギャップが、元VBA・前回DLLのログ記録精度で一致。
- 保存した入力セル・節点データ・要素データが、元VBA・前回DLLの保存結果と一致。
- 前回のDouble/DD/Schur数値核も実Excelで再検証し、解・hi/lo・最小ピボットが8バイト単位で一致。
- DLLなしの別コピー・別Excelで、数値LDL/Schurと両並べ替え経路のVBA自動切替を確認。

ログ・保存セルの一致は、実解析の全内部状態のビット単位同一性の証明ではない。
今回の実解析は堤防1例。25ケース全回帰、DDを多用する実解析・非関連流れ則ケースの高速化率は未検証。
堤防での呼出し回数：`{final['native_calls']}`。

## 中断・異常時の扱い

グラフは呼出し単位の所有コンテキストで管理し、VBAの配列ポインターを呼出し後へ保管しない。
入力転送256行、消去64行ごとにVBAへ戻り、DoEvents・取消を確認する。
消去中は入力同一性も確認する。元の8192行ごとの進捗表示を保持した。
通常終了・例外・取消のすべてでグラフコンテキストを解放する。
1回のDLL呼出しの処理中はVBAへ戻らないため、極端に密な行や初期化時の応答時間は行列規模に依存する。

| 試験・防御 | 結果・区別 |
|---|---|
| 全体コンパイル | 実ExcelでPASS |
| 取消注入 | 実Excelでエラー18を保持。数値核・並べ替えで確認 |
| 範囲外列・短いバッファ | 実Excelでエラー9。ネイティブへ不正ポインターを渡さない |
| 特異ピボット・桁あふれ | 実Excelで元VBAと同じFalse・エラー6を確認 |
| メモリ確保失敗 | C++単体試験で**2835か所**へ順次注入。戻りコード207、解放後の未解放確保0。VBA側はエラー7へ伝達する実装 |
| 未知例外 | C++単体試験で8か所へ注入。戻りコード208、解放確認。VBA側はエラー513へ伝達する実装 |
| 途中区間で終了 | C++単体試験で解放後の未解放確保0 |
| 不正契約 | C++単体試験で未初期化・列範囲・短い出力・不正区間を拒否 |
| 原本保全 | 元G1R15の152ファイルと、前回DLL成果物・ZIPを含む171ファイルのハッシュ一致 |
| 旧P2S基準保全 | verify_baseline.pyで32ファイル一致。旧P2SのExcel試験はNOT_RUNのまま |

注入用メモリ確保器は試験EXEだけに存在し、配布DLLへ含めていない。
メモリ確保失敗・未知例外の注入はC++単体試験であり、実Excelでメモリを枯渇させる試験は実施していない。
失敗を解析成功や通常のリメッシュ失敗として隠さず、既存のエラー分類へ戻す。
DLL不在・読込不可・export不足・ABI不一致の場合はVBAを使う。ロード後の任意の異常を一律にVBA再試行へ変更しない。

## 時間の内訳

| 区分 | 前回DLL | DLL2 |
|---|---:|---:|
{phase_table}

一部の区分は入れ子なので単純合計しない。
並べ替えの短縮分がすべて総時間の短縮にはならず、今回は疎行列への変換時間が増えた。
VBAの行オブジェクト再利用とネイティブ処理の関係など、変換段階の増加理由は今回まだ特定していない。
残った線形求解・補正には、VBAの残差計算や品質確認を含む。
今回の結果には、未実装のNT・残差・行列積の移植効果を含めていない。

## 条件とファイル

- 元コード：GitHub main `d6c4d0903582a325d7c0368cf22dbfb80ee64db6` のG1R15。
- 元ブック：`versions/G1R15/examples/RPFEM_20261011_G1R15_Levee_Example.xlsm`。
- 堤防：高さ5m、天端幅4m、基礎を含む幅40m・最大高さ13m。c=10kPa、φ=ψ=25度、γ=19kN/m³。
- 浸透連成ON、γw=9.81、サクションOFF、kh=0.15右(+x)。
- TARGET=600、ADAPT=1、CYCLES=3、QUADRATURE=8、GAP=0.01、FS_TOL=0.0001。
- AMD Ryzen 9 5900X、64bit Excel 16.0。同じPCで実行。

配布XLSM：`delivery/RPFEM_20261012_G1R15_DLL2_Levee_Checked.xlsm`。解析結果・条件を保存済み。
DLL：`delivery/rpfem_native_x64.dll`。必ずこの版を同じフォルダーへ置く。
編集用UTF-8は `src/vba/`、VBE用の無損失CP932/CRLF標準モジュール6件は `import/`。
元ソース・差分・試験・ビルド用ソース・ログ・JSON照合・ハッシュ一覧を同梱した。
文書/クラスモジュールは変更していない。原本を変更せず、ネイティブExcelで別コピーを更新・保存した。
vbaProject.bin差替え、ライター再構築、信頼設定・マクロセキュリティ変更、GitHub公開は行っていない。

LDL派生コード・DLLはLGPL-2.1-or-later。追加並べ替えソースは元プロジェクトと同じApache-2.0。
対応するライセンスと再ビルド可能なソースを `native/` に同梱した。

XLSM SHA-256：`{verification['xlsm_sha256']}`

DLL SHA-256：`{verification['dll_sha256']}`
'''
(ROOT/'REPORT.md').write_text(report,encoding='utf-8')
(ROOT/'README.md').write_text(f'''# G1R15 DLL2・堤防例題

前回DLL **{comparison['previous_dll_seconds']:.2f}秒 → {final['analysis_seconds']:.2f}秒**。
追加で **{final['reduction_percent_vs_previous_dll']:.2f}%短縮**。元VBAの既存622.12秒からは **{final['reduction_percent']:.2f}%短縮・{final['speedup']:.3f}倍**。
Fs下界1.224448、上界1.234308、監査603件PASS、同じ53求解。

## 使い方

1. ZIPを展開する。
2. `delivery/` のXLSMとこのZIPのDLLを同じフォルダーへ置く。
3. 現在の許可済みExcel設定で開き、通常どおり解析する。

DLLなしの場合はVBAへ自動切替する。64bit Excelで検証済み。
前回版と同じDLL名のため、前回の解析用Excelを閉じてから新しい版を開く。
ロード済みのDLLはそのExcel終了まで置き換えず、フォルダーを分けて利用する。
旧DLLと新XLSMを混ぜた場合は、export/ABIチェックによりVBAへ戻る。

条件：堤防例題、浸透連成ON、kh=0.15右、TARGET=600、メッシュ最適化ON。
比較の範囲と試験結果は [REPORT.md](REPORT.md)。導入・復旧は [IMPORT_ROLLBACK_JA.md](IMPORT_ROLLBACK_JA.md)。
''',encoding='utf-8')
(ROOT/'IMPORT_ROLLBACK_JA.md').write_text('''# 導入・復旧

配布XLSMは導入済み。ZIP内のXLSMとDLLを同じフォルダーへ置く。
前回DLL版とファイル名が同じDLLを使うため、前回版を開いていたExcelを終了してから開く。
DLLがない場合は、数値核と並べ替えが既存VBAへ自動切替する。

## 元G1R15への手動導入

原本を保全し、別コピーへ導入する。現在の許可済みVBE設定を使用する。

1. RPX_LDL、RPX_SchurLDL、RPX_SeepUI、RPX_Order、RPX_ColumnOrderをエクスポートして保全する。
2. その5標準モジュールを削除し、同名の `import/*.bas` を導入する。
3. 新規標準モジュール `import/RPX_Native.bas` を導入する。
4. 全体をコンパイルして別名保存し、このZIPのDLLを同じフォルダーへ置く。

前回DLL版へ導入する場合も別コピーを使う。RPX_Nativeは追加でなく置換する。
`import/` はCP932・CRLF、往復無損失検証済み。`src/vba/` は編集用UTF-8で、直接VBEインポートしない。
ThisWorkbook・Sheet・クラスモジュールへの配置変更は不要。それらを標準モジュールとして導入しない。

## 復旧

完全復旧は保全した原本、または前回成果物へ戻す。
導入コピーの既存5モジュールを原本へ戻してRPX_Nativeを削除する方法でも元G1R15へ復旧できる。
DLLだけ外す場合はVBAへ戻るが、呼出しフックと前回のプレビュー関数名修正は残る。
DLLを外す・置き換える操作は、そのDLLを利用しているExcelを終了してから行う。

## 再ビルド・試験

`native/build.cmd` で64bit DLLを生成する。Visual Studio Build Toolsのvcvars64.batのパスは環境に合わせる。
数値核のx87式評価・保存境界・演算順を変えない。今回は整数の並べ替えを追加した。
`tests/build_failure_test.cmd` はメモリ確保失敗と未知例外を注入する専用EXEを作り、製品DLLを変更しない。
`tests/NativeTests.bas` は実Excelでの比較用。配布XLSMには残していない。

Pythonツールはpywin32等を必要とする。通常のXLSM解析にPythonは不要。
編集後は `tools/prepare_imports.py` でCP932ファイルを作成する。
旧 `prepare_sources.py` は編集済みソースを初期化するため、この版では実行を拒否する。
`install_and_test.py` は保全済み元G1R15の別コピーへ導入・コンパイル・短い比較試験を行う。
`run_benchmark.py` は別フォルダーで解析し、既存Excelを強制終了しない。
比較・原本ハッシュ検証の一部は、元リポジトリと前回成果物のフォルダー配置を参照する。
実行例・結果は `results/`。ハッシュ一覧はMANIFEST.jsonとSHA256SUMS.txt。
''',encoding='utf-8')
reference=ROOT/'results/reference_logs'; reference.mkdir(exist_ok=True)
shutil.copyfile(ROOT.parent/'G1R15/results/RPFEM_logs/20261011_211428_26004_analysis.tsv',reference/'original_vba_622s.tsv')
shutil.copyfile(next((ROOT.parent/'G1R15_DLL/results/dll_run3_final/RPFEM_logs').glob('*analysis.tsv')),reference/'previous_dll_444s.tsv')
excluded={'.obj','.lib','.exp','.exe','.pyc'}; files=[]
for p in ROOT.rglob('*'):
    if not p.is_file() or p.suffix in excluded or '__pycache__' in p.parts:continue
    rel=p.relative_to(ROOT)
    if rel.parts[:2] in [('tests','missing_dll'),('tests','missing_order')]:continue
    if p.name.startswith(('template_before_','compile_failed_')):continue
    if rel.parts[:2]==('results','dll_run1_ordering') and p.suffix in {'.dll','.xlsm'}:continue
    if p.name in {'MANIFEST.json','SHA256SUMS.txt'}:continue
    if p.name=='finalize.py':continue # Superseded by this script.
    files.append(p)
manifest={str(p.relative_to(ROOT)).replace('\\','/'):hashlib.sha256(p.read_bytes()).hexdigest() for p in files}
(ROOT/'MANIFEST.json').write_text(json.dumps(manifest,ensure_ascii=False,indent=2),encoding='utf-8')
(ROOT/'SHA256SUMS.txt').write_text(''.join(h+'  '+name+'\n' for name,h in sorted(manifest.items())),encoding='utf-8')
files += [ROOT/'MANIFEST.json',ROOT/'SHA256SUMS.txt']
archive=ROOT.parent.parent/'RPFEM_G1R15_DLL2_Levee_Checked.zip'
with zipfile.ZipFile(archive,'w',zipfile.ZIP_DEFLATED,compresslevel=6) as z:
    for p in files:z.write(p,'G1R15_DLL2/'+str(p.relative_to(ROOT)).replace('\\','/'))
with zipfile.ZipFile(archive) as z:
    assert z.testzip() is None
    for name,h in manifest.items():assert hashlib.sha256(z.read('G1R15_DLL2/'+name)).hexdigest()==h
print('PACKAGE PASS',archive,'files',len(files),'bytes',archive.stat().st_size)
print('FINAL',final['analysis_seconds'],'vs previous reduction',final['reduction_percent_vs_previous_dll'],'vs original speedup',final['speedup'])
