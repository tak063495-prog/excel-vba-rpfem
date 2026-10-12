"""Generate Japanese report from executed-test evidence, not inferred passes."""
import csv,json
from prepare import ROOT,write,sha
def read(p):return json.loads((ROOT/'results'/p).read_text(encoding='utf-8'))
def main():
 c=read('comparison.json');v=read('delivery_verification.json');roots=read('root_regressions.json');bench=read('fixed_q8_threeway/summary.json');bulk=read('bulk_oracle/summary.json');lu=read('lu_oracle/summary.json');nt=read('nt_oracle_extended/comparison.json');boundary=read('boundary_oracle/summary.json');contract=read('job_contract/summary.json');lifecycle=read('job_lifecycle.json');issue=read('issue_regression/summary.json');hsd=read('hsd_regression/summary.json');fallback=read('fallback_seep.json');faults=read('native_faults.json')
 dll=sha(ROOT/'delivery/rpfem_native_x64.dll')
 assert all(r['dll_sha256']==dll for r in [c,v,bench['backends']['dll4'],bulk,lu,nt,boundary,contract,lifecycle])
 assert bench['passed'] and contract['passed']==12 and issue['failed']==0 and issue['passed']==51
 assert bulk['exact'] and not lu['failures'] and not nt['differences'] and not nt['classification_mismatches'] and boundary['bitwise_exact']
 assert len(hsd['cases'])==6 and all(r['bitwise_exact'] for r in hsd['cases']) and fallback['same_hsd_result_all_backends'] and all(r['exit_code']==0 for r in faults.values())
 phase_rows='\n'.join(f"| {p} | {c['DLL3_phases'].get(p,0):.3f} | {c['DLL4_phases'].get(p,0):.3f} |" for p in ['nt_scaling','kkt_update','corrector','gproduct','gtadd','linear_solve_and_refinement','ordering'])
 fixed_rows='\n'.join(f"| {key} | {r['dll4']['elements']} | {r['dll4']['Fs_input']:g} | {r['dll3']['seconds']:.3f} | {r['dll4']['seconds']:.3f} | {r['reduction_percent_vs_DLL3']:.2f}% | ビット一致・監査合格 |" for key,r in bench['cases'].items())
 def fs_text(r,side):return f"{r['DLL4'][side]:.15g}" if r['DLL4']['Fs_obtained'] else '未取得'
 root_rows='\n'.join(f"| {r['case']} | {'終端まで' if r['full_case_compared'] else 'Fs≥0.125・15点'} | {fs_text(r,'lower')} | {fs_text(r,'upper')} | {r['DLL4']['terminal_status']} | {r['comparison_seconds']['DLL3']:.2f} | {r['comparison_seconds']['DLL4']:.2f} | {r['reduction_percent_vs_DLL3']:.2f}% |" for r in roots['paired_roots'])
 failed_roots=[r['case'] for r in roots['paired_roots'] if not r['case_passed'] and r['full_case_compared']]
 partial_roots=[r['case'] for r in roots['paired_roots'] if r['DLL4']['Fs_obtained'] and not r['DLL4']['target_met']]
 assert all(r['comparison_passed'] for r in roots['paired_roots'])
 p06=roots['P06A_full_mesh']
 text=f'''# G1R16 DLL4 検証報告

2026年10月12日、同一PCの実Excel 16.0・64bitで検証しました。通常内点法とHSDの数値演算を追加DLL化し、線形求解・精度補正を状態付きの処理としてDLL内で進めます。

## 堤防例題の速度と精度

**DLL_3 {c['DLL3_seconds']:.2f}秒 → DLL4 {c['DLL4_seconds']:.2f}秒、{c['reduction_percent_vs_DLL3']:.2f}%短縮、約{c['speedup_vs_DLL3']:.2f}倍**でした。

| 項目 | DLL_3（今回の再実行） | DLL4 |
|---|---:|---:|
| 解析時間・ログ（秒） | {c['DLL3_seconds']:.6f} | {c['DLL4_seconds']:.6f} |
| COM呼出し全体（秒） | {c['DLL3_COM_seconds']:.6f} | {c['DLL4_COM_seconds']:.6f} |
| Fs下界 | {c['lower']:.17g} | {c['lower']:.17g} |
| Fs上界 | {c['upper']:.17g} | {c['upper']:.17g} |
| 求解 | {c['DLL3_log_comparison']['solves']} | {c['DLL3_log_comparison']['solves']} |
| 監査記録 | {c['DLL3_log_comparison']['audits']} | {c['DLL3_log_comparison']['audits']} |

TARGET=600、QUADRATURE=8、適応ON、kh=0.15右、浸透連成ONの同じ入力・同じメッシュです。先にDLL_3、次にDLL4を新しいExcelプロセスで実行し、他の本タスクの長時間計算とは重ねていません。表は`RPX_Run`の`run_end`までです。事前の浸透計算・メッシュ生成、ファイル保存、Python監査は含みません。全解析の比較を、線形核だけの時間と混同していません。単一PC・各1回の比較で、どのPCでも同じ改善率になるという意味ではありません。

Fsは保存済みDoubleとしてビット一致。全求解の探索順・Fs点・反復数・目的値・残差・GAP、全物理監査の詳細はDLL_3および実G1R16 VBAの既存ログと記録精度で一致しました。入力・節点・要素の保存済みデータも一致しています。VBAの長時間速度測定は再実行していません。既存VBAログの{c['VBA_existing_seconds']:.2f}秒は参考値です。

| 相別の処理 | DLL_3 秒 | DLL4 秒 |
|---|---:|---:|
{phase_rows}

相別の記録は範囲が重なるため、合計は全解析時間と一致しません。演算をまとめたことで相への計上先も変わるため、個別の相の時間だけから演算の短縮量を判断しないでください。`results/comparison.json`に生の時間・比較・証跡SHA-256があります。

## VBA・DLL_3との数値比較

| 試験 | 実測した結果 | 証跡 |
|---|---|---|
| HSDの既知SOCP解 | LDL_CHECKED・PIVOT_ONLY計6組。VBA・実DLL_3記録・DLL4の解、反復数、残差、DD経路を含む出力がビット一致 | `hsd_regression/`、`three_way_precision.json` |
| LU | 360条件、352の因子・ピボット・解が完全一致、特異8も同じ拒否 | `lu_oracle/` |
| LUの独立残差 | Pythonで最大後退誤差{lu['max_independent_backward_error']:.3e} | 同上 |
| NT | 224入力、成功174はDouble・DD hi/lo完全一致、拒否50は例外判定一致 | `nt_oracle_extended/` |
| 一括NT・補正・KKT・G/G転置・DD原作用素 | 160組、{bulk['bytes_vba']:,}バイト完全一致 | `bulk_oracle/` |
| コーン境界・内部性 | 5,070条件。極端な値とオーバーフロー時の途中出力・例外も一致 | `boundary_oracle/` |
| LDL・DD・Schurと並べ替え | 数値核・特異・短い配列・オーバーフロー・取消合格。順列/逆順列68組一致、n>20000を含む | `bulk_oracle/summary.json` |

固定FsのFEMは最終VBAソース・DLLで、TARGET=256・q=8・適応OFF・浸透OFF・kh=0として各版を実行しました。配布XLSMは、その同じソースで堤防の結果を実Excelに保存したブックです。Fs欄は**求めた安全率ではなく入力した固定点**です。C5_45はc=5、φ=ψ=45。P06A/P06Nは多層・関連/Davisモデルです。VBA、DLL_3、DLL4でメッシュと全場のバイナリが一致し、Pythonで独立物理監査を実行しました。

| 固定点・側 | 実要素数 | 入力Fs | DLL_3 秒 | DLL4 秒 | 短縮 | 3版の比較 |
|---|---:|---:|---:|---:|---:|---|
{fixed_rows}

各固定点1回、試験関数`FixedFem`のCOM呼出し全体の時間です。入力読込み・メッシュ生成・求解・場データの出力を含み、Excelの起動とPython監査は含みません。堤防全解析の時間・短縮率とは区別します。q=2・TARGET=64でも6組と独立監査を実行していますが、それは最終の時計・メモリ見積り・切替ガード修正前の開発時証跡です。最終版の比較根拠は上記q=8です。

## Fs探索の回帰と難しいP06

P06A・P06N・D03N・D07Aは元のケースの形状・材料・境界を使用し、TARGET=200、q=8、適応OFFで比較しました。同じ入力・メッシュ、比較範囲内の求解・監査ログの一致を確認しました。P06N・D03N・D07Aは`RPX_Run`の終端まで比較し、Fsを取得できたケースは最終Fsがビット一致しています。P06Aの追加比較は、下記の15点に限定しています。

| ケース | 比較範囲 | Fs下界 | Fs上界 | DLL4判定 | DLL_3 秒 | DLL4 秒 | 短縮 |
|---|---|---:|---:|---|---:|---:|---:|
{root_rows}

P06N/D03N/D07Aの時間は`RPX_Run`のCOM呼出しの実測です。P06Aは後述の15点目の完了ログ時点を使っています。Fsを得ても目標区間幅に未達だったケースは **{', '.join(partial_roots) if partial_roots else 'なし'}** です。これはFs取得の比較合格と目標達成の区別です。

**粗メッシュの終端比較でFsを取得できなかったケース：{', '.join(failed_roots) if failed_roots else 'なし'}。** 表の未取得を安全率0や解析合格とは扱いません。失敗ケースでは両版とも各固定Fsの求解は進みますが、下界荷重係数が1を越える点を囲めず、元のエラー5 `FS_LOWER_BRACKET_NOT_FOUND`に到達します。全探索点・反復数・目的値・残差・監査と終端エラーまで一致することを確認し、「基準版との回帰一致」と「安全率取得の合格」を分けました。これは粗メッシュ条件での未達であり、フルメッシュでも求解不能と断定する結果ではありません。メッシュ依存かモデル・境界・非関連流れ則の制約かという原因の確定は今回のDLL化の精度比較だけではできません。

**P06Aの追加比較は15点だけです。** 基準DLL_3の粗メッシュ探索は約28.87分、23求解目のFs=0.00048828125で150反復の上限に達し、`HSD_NOT_CONVERGED`となりました。極小Fsの補正に時間を要したため、検証範囲の選択肢を提示しました。ユーザーが「極小Fs探索を終了し、取得済みの実用Fs点をDLL4と照合して仕上げる」を選択したため、取得済みのFs≥0.125の15点へ絞りました。中断マクロを呼ぶ前に基準計算は終了しており、取消スクリプトは対象の実行終了を検知して何も変更していません。

DLL4は試験コピーの`FsValue`入口だけにFs<0.125で取消18を送る処理を追加しました。比較した15点の演算、元の許容差、補正・救済、製品のFs探索コードは変更していません。試験コピーの変更がその停止処理だけであることも確認しました。表のP06A時間は**同じ15点を終えたログ時点**の比較で、基準版の約28.87分と短い試験時間を直接比べていません。Fs取得、全粗メッシュ探索の終端一致、23点目の未収束のDLL4再現は検証完了としていません。

失敗ケースの時間は終端エラーまでの実測です。安全率の完走時間やフル適応計算の短縮率には換算しません。

初期の比較ハーネスは全粗メッシュで正のFsを要求して停止しました。その原記録は`pipeline_initial_failure.json`に保全し、既存の名前付き下界ブラケットエラーだけを失敗ケースとして記録する比較へ変更しました。取消・メモリ不足・添字・監査・未知例外はこの扱いに含みません。製品コード、メッシュ、物理モデル、許容差を変えて失敗を隠す変更はしていません。

P06AのTARGET=1536・q=8・適応ON（設定CYCLES=3）は開始しましたが、**{p06['seconds']/60:.1f}分でユーザー指示により中断**しました。既存の`RPX_Stop`で取消18を伝え、実ログは **{p06['terminal_status']}** です。プロセスの強制終了はしていません。最終Fsが出ていないため、完走・精度合格・目標達成・全探索の速度比較として数えていません。P06Nのフル適応解析も開始していません。両ケースの比較は上記の固定Fsと粗メッシュFs探索です。長時間試験のブックは切替時の因子無効化修正前に作成しましたが、DLLは最終版です。

`PARTIAL_CERTIFIED`は監査済みの上下界を得たという意味で、目標GAPやFs探索幅の合格ではありません。採用した端点と目標達成を区別しています。今回のDLL化は離散化に由来する上下界差を縮める変更ではありません。

カタログJSONの`GRAVITY`は元から「固定」で、Git基準版とバイト一致を確認しました。両版へ同じ値を明示設定しています。開始時の試験メタデータにあった「文字化け」という診断は表示の誤認であり、元データの変更はありません。複雑な回帰の`HSD_SOLVE_SECONDS=7200`は試験だけの時間上限です。配布堤防ブックの標準値は変更していません。今回は26ケースの全再実行や、P06N/D03N/D07AのTARGET1536適応全探索を合格扱いしていません。

## ロバスト性

- 通常/HSDそれぞれでジョブ開始後に、取消18、メモリ不足7、添字9、未知例外513、イベント無効下の入力変更5、通常完了を試験。12件すべて番号・理由を保持し、全ハンドルが0へ戻りました。
- 262,144変数の通常/HSDジョブで実際の途中返却を確認。別ジョブとの交互進行、途中解放、異なるepoch・因子・早すぎる結果取得・二重開始・解放済みIDを試験し、解が混ざらず、残るハンドルは0でした。
- バックエンド切替後にVBA側だけ因子readyが残る不整合を再現し修正。auto→vba、auto→autoとも因子を両側で無効化し、新しい分解後の求解が合格しました。
- C++故障注入の実ビルド/実行：並べ替え2,835箇所、LU66箇所、通常ジョブ6箇所、HSDジョブ15箇所の割当失敗と、未知例外の注入が合格。エラー207/208を返し、解放漏れなし。VBAではそれぞれ7/513へ明示変換します。
- Issue #1/#2の51回帰、DLL不在・旧DLL2・旧DLL_3のVBA切替、浸透収束/未収束・水位入力の回帰が合格しました。

状態の所有、入力確認、元の補正・救済規則を維持しています。DLL途中の未知例外をVBA再実行で隠しません。約50msの作業区切りとDoEventsを維持していますが、不可分処理があるため最大応答時間の保証ではありません。製品は単一Excelスレッドで使用し、並列C API使用は検証していません。

## 保存・導入・再現

元版533ファイルのSHA-256不変。既存68モジュールから6を変更し1を追加、最終69モジュール。62既存モジュールと文書モジュールは不変です。UTF-8編集ソース、無損失CP932/CRLF取込み、保存ブックとのソース一致、実Excelの全プロジェクトコンパイル・再起動を確認しました。元のXLSMは保存し、別名コピーを実Excelで更新しています。VBAバイナリの入替、マクロブックのライター再構築、セキュリティ・レジストリ変更はしていません。

開始時の`baseline.json`にDLL2を指すラベルが残っていました。実際の68モジュールの基準と導入元はDLL_3で、`install.json`の元SHA-256と保存ソース比較で確認しました。ラベルを訂正し、元の記録と訂正理由は`initial_metadata_discrepancy`に残しています。

ABI `0x00160004`。XLSMとDLLを組で同じフォルダーへ置き、交換時はExcelを終了します。詳細は[導入・復旧](IMPORT_ROLLBACK_JA.md)と[実装範囲](IMPLEMENTATION_JA.md)。線形演算のまとまりはDLL内で進みますが、固定Fsの非線形反復制御・Fs探索・監査・救済・メッシュは既存VBAにあります。

Pythonの模擬試験、C++ API試験、静的保全、実Excelコンパイル、実VBA数値比較、実Fs探索の結果をそれぞれ区別しています。証跡は`results/`、再現コードは`tests/`と`tools/`、差分は`changes.diff`。再現にはリポジトリ全体、64bit Excel、pywin32/numpy/scipy/oletools、VS Build Toolsが必要です。検証用コードを製品ブックへ追加する処理は各試験コピーに限定します。

DLL SHA-256: `{dll}`

XLSM SHA-256: `{v['xlsm_sha256']}`
'''
 (ROOT/'REPORT.md').write_text(text,encoding='utf-8',newline='\n')
 readme=f'''# G1R16 DLL4

通常内点法とHSDの線形求解・精度補正・残差・方向・境界演算を追加DLL化した64bit Excel版です。堤防例題は **DLL_3 {c['DLL3_seconds']:.2f}秒 → DLL4 {c['DLL4_seconds']:.2f}秒（{c['reduction_percent_vs_DLL3']:.2f}%短縮）**。Fs下界・上界と全求解・監査がVBA/DLL_3の比較に合格しました。

1. Excelを終了し、ZIPを新しいフォルダーに展開します。
2. `delivery/`の対応XLSMとDLLを同じフォルダーに置きます。
3. マクロを通常の手順で実行できる状態にし、XLSMを開きます。
4. 初期入力は浸透・地震を含む堤防例題。「要素作成②」から入力・メッシュを確認し、「解析」を実行します。
5. `RPX_NativeInfo`の`DLL_READY`でDLL使用を確認できます。DLL不在・非対応版ではVBAへ切り替わります。

- [検証報告と速度表](REPORT.md)
- [実装範囲・精度・状態の寿命](IMPLEMENTATION_JA.md)
- [取込みと復旧](IMPORT_ROLLBACK_JA.md)
- [VBA差分](changes.diff)、[編集ソース](src/vba/)、[取込みファイル](import/)、[DLLソースとライセンス](native/)

XLSMとDLLは組で使用します。DLLだけの追加ではDLL4になりません。構成則、Davis変換、荷重、Fs定義、許容差、物理監査を維持しています。Fs取得と目標GAP達成は別で、`PARTIAL_CERTIFIED`を目標達成として扱わないでください。

P06の3版q8固定点・独立物理監査、P06N/D03N/D07Aの粗メッシュ終端比較、P06AのFs≥0.125の15点比較、HSD/LU/境界と異常時試験を実施しました。粗メッシュの一部では基準DLL_3でもFsを取得できず、失敗結果と比較合格を区別しています。P06Aフル適応解析はユーザー指示で中断し、P06Nフル適応解析は開始していません。P06A粗メッシュの極小Fs未収束をDLL4で再現する試験も完了していません。26ケースすべての再実行ではありません。詳しい条件・未実施範囲は報告を参照してください。

ビルドは`native/build.cmd`。VS Build Toolsの実際のパスに合わせます。通常の解析にPythonやVBAプロジェクトアクセス許可は不要です。検証スクリプトの実行環境は`tools/REPRODUCTION_JA.md`を参照してください。
'''
 (ROOT/'README.md').write_text(readme,encoding='utf-8',newline='\n')
 rows=[('Original533','static','PASS','533 immutable hashes','results/delivery_verification.json'),('ExcelCompile','actual Excel','PASS','69 modules; reopen/compile','results/delivery_verification.json'),('HSD_threeway','actual VBA/DLL3/DLL4','PASS','6 canonical SOCP pairs','results/three_way_precision.json'),('FEM_q8_threeway','actual VBA/DLL3/DLL4 + independent Python audit','PASS','6 fixed-Fs pairs; TARGET256 q8','results/fixed_q8_threeway/summary.json'),('LU','actual VBA/DLL4 + independent Python residual','PASS','360;352 exact+8 singular','results/lu_oracle/summary.json'),('NT','actual VBA + native API','PASS','224;174 exact+50 matching rejections','results/nt_oracle_extended/comparison.json'),('Bulk','actual VBA/DLL4','PASS','160;1239592 bytes exact','results/bulk_oracle/summary.json'),('Boundary','actual VBA/DLL4','PASS','5070 exact including partial outputs on overflow','results/boundary_oracle/summary.json'),('NativeFaults','native C++ test builds and runs','PASS','allocation/unknown exception injection; no leaks','results/native_faults.json'),('JobControls','actual VBA/DLL4 failure injection','PASS','12; cancellation/OOM/subscript/unknown/input/success','results/job_contract/summary.json'),('JobLifecycle','native API','PASS','262144 variables; real yields/interleaving/stale guards','results/job_lifecycle.json'),('ModeSwitch','actual VBA/DLL4','PASS','2 modes invalidate stale factor; new solves','results/mode_switch_fixed/summary.json'),('Issue1_2','actual Excel','PASS','51 assertions','results/issue_regression/summary.json'),('Fallback_Seep','actual Excel','PASS','missing/old DLL2/old DLL3; seep convergence regressions','results/fallback_seep.json'),('Levee','actual complete RPX_Run DLL3/DLL4; existing VBA reference','PASS','same Fs bits; complete solve and audit log matches','results/comparison.json')]
 rows.extend((r['case']+'_root','actual RPX_Run DLL3/DLL4',r['DLL4']['terminal_status'],'TARGET200 q8 ADAPT0; '+('Fs bits and logs exact' if r['case_passed'] else '15-point bounded test; no final Fs' if not r['full_case_compared'] else 'Fs unavailable; named baseline bracket failure and logs match'),'results/root_regressions.json') for r in roots['paired_roots'])
 rows.extend((r['case']+'_root_baseline_comparison','actual RPX_Run DLL3/DLL4','PASS','Same solve/audit sequence and '+('final outcome' if r['full_case_compared'] else '15 completed Fs points only; final outcomes not compared'),'results/root_regressions.json') for r in roots['paired_roots'])
 rows.append(('P06A_extreme_failure_reproduction','actual RPX_Run DLL4','NOT_RUN','DLL3 failed at Fs0.00048828125; DLL4 comparison bounded above that point','results/root_regressions.json'))
 rows.append(('P06A_adaptive','actual RPX_Run DLL4',p06['terminal_status'],'TARGET1536 q8 ADAPT1; no same-size DLL3 comparator','results/root_regressions.json'))
 rows.append(('P06N_adaptive','actual RPX_Run DLL4','NOT_RUN','Full adaptive search not started per user-selected comparison scope','results/root_regressions.json'))
 with (ROOT/'TEST_RESULTS.csv').open('w',encoding='utf-8-sig',newline='') as f:
  writer=csv.writer(f);writer.writerow(['test','validation_type','result','executed_scope','evidence']);writer.writerows(rows)
 print('REPORT evidence assertions PASS',flush=True)
if __name__=='__main__':main()
