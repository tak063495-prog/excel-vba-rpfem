# D02A: Pythonと配布G1R19 VBAの経路監査

2026-10-16。ユーザー報告のD02A TIMEOUTを受けた調査。
D02Aの正本ログを [PR #20](https://github.com/tak063495-prog/excel-vba-rpfem/pull/20)、main `303922050bea7ddaaf6f26fe560c1a4d10c76121` の `validation/G1R19_preset2pct_20261016/D02A/` から受領した。前回のAPI検索後に公開された新着である。生バイトをGit blob SHAと照合し、UTF-16LE原ログとUTF-8版の復号文字列一致を確認した。

## 現時点の結論

Pythonの成功は、配布G1R19のVBA制御で同じ結果が自動的に再現される証拠ではなかった。
物理式・厳格な場の監査を共有していても、候補生成の順序・メッシュ・q・求解器の選択が異なる。
とくにD01A/D02Aのような全c=0斜面で、Python側が先に使う有効な上限候補が
配布VBAの既定設定では無効であり、速度プリセットでも初段の直接DLL経路が使われない。

## 確認できた差分

|項目|新規Python検証|配布G1R19 VBA|影響・対応|
|---|---|---|---|
|初段の次数・根幅|最初からq8、正規化された根幅1e-4|速度プリセットが`ADAPT_PRECISION_FROM_START=0`。初めの64求解はq4・根幅1e-3。852秒からfinal q8へ切替|最初のq4段では直接求解のq8条件を満たさずHSD経路。final以降はDirect DLL144回、117回採用・27回fallbackを実測。最初から無効だったとは扱わない|
|c=0上限候補|モデルの適合性を確認し、斜面高さの0.001倍の薄層8要素候補を最初に求解。元q8式と自己メッシュ上の根を再確認|`ADAPT_C0_WEDGE_DEPTH=0`で候補は無効。呼出位置も適応計算の後|適合条件を保持した自動候補を早期に呼び出す修正を実装中。解析式の値そのものを上限としない|
|細分化指標|速度ひずみと内部ジャンプに基づき、c=0でも非zero|LEGACYのせん断指標はcを掛けるため全c=0で0になり得る|0の場合のみ既存の運動学的指標へ切替。D02A実ログの294.70秒でband moved=0/weight=0、その後band候補は悪化してrollback。後半accuracy段は既にkinematic指標を使用|
|下限端点の証拠|`NumericalError`等でも原式の物理場・負荷余裕・root条件を独立に確認。status単独では採否を決めない|HSD UNBOUNDEDのrayをλ=1へ正規化する。原場監査を通ったrayでも、`UNBOUNDED`のままでは`RPX_ActualRootBracket`が要求する端点status区分を満たさない経路がある|原b/h=0等を厳格に確認し、同Fsの有限rayコピーをλ=1.000003へ変換後、原式を再監査して既存のphysical-witness区分で扱う。単なるstatus書換えはしない|
|根の補間|物理端点と原dual guideで符号が確認された区間を収束させる|Directの正側primalは任意の負荷量λ=2.34～25629.8等、負側は原dualによる上限であり、同じ最適値関数の標本ではない。旧HSD UNBOUNDEDのRValue=1e30はFsValue末尾で1へ置換されるため、1e30が必ずsecantへ入るとは言わない|この組合せは符号に基づく二分法に切替。根の受入条件・幅は維持|
|初期メッシュ|G1R17検証に保存された初期形状に対して新規求解、角度fanを適用|現VBAのTARGET=1000等から初期メッシュを作成|入力物理モデルが同じでも行列・メッシュは同一でない。Python時間をVBAや別PCとの厳密な速度比として報告しない|
|LPの表現|`Prepared(..., reduction='affine')`で等式を消去した候補求解を使い、復元した場・双対を原全空間で監査|直接DLL経路は原全空間LPで、AFActiveの場合は対象外|求解サイズ・条件数も異なり、同じClarabelというだけでは同じ処理にならない|
|内部solver設定とfallback|各LP上限30秒。最初は内部tol1e-10、必要時に反復改良・1e-13・cone scaling・equilibrationへ切替。原場/双対の採用閾値は不変|直接LP上限の既定5秒・内部tol1e-13・反復改良30回。失敗時HSDと別回復へ移る|最初の候補求解の条件・失敗時の費用も違う。内部tolと最終的な原物理監査の閾値を区別して測定する|

表は公開ソースの経路とD02A原ログで観測した事実を区別して記す。rayの監査合格や補間の片寄りが全ての求解で必ず起こるとは限らず、停止後の無応答の正確な命令位置はログから確定できない。

原VBAの根拠:

- `versions/G1R19/src/vba/RPX_UI.bas`: `RPX_Speed2pctPreset`、`ADAPT_C0_WEDGE_DEPTH`既定0。
- `versions/G1R19/src/vba/RPX_Adapt.bas`: `Evaluate`粗段q4制御、`ComputeShearWeight`、`ContinueUpperC0Wedge`、適応終了後の呼出。
- `versions/G1R19/src/vba/RPX_AccuracyNative.bas`: `RPX_DirectEnabled`は`RSubdivisions=8`を要求。
- `versions/G1R19/src/vba/RPX_Fs.bas`: c=0でHSD選択、UNBOUNDED rayのλ=1正規化、`RPX_ActualRootBracket`の証拠条件。
- `validation/G1R19_python26_20261016/python/accuracy_driver.py`: `search`のLP候補群、原dual/Farkasの再監査、固定負荷候補の採用区分。
- `validation/G1R19_python26_20261016/python/experiment_core.py`: `Prepared`、`settings`、原全空間へ復元した場の監査。

Pythonパッケージと単一DLLの新しいaccuracy bridgeはどちらもClarabel 0.11.1だが、
VBAの従来HSD求解はこの直接bridgeとは異なる。
言語処理速度だけの差として扱うことはできない。

## 受領したD02A TIMEOUTログの確定事項

公式ブックSHAは `3533073bfe883b254143cf68285235ddbe729e8f0e272f4180f2b730f9c63a6f`、DLLは `66da1a21be43cd2b3a3bfbaa9d13aad5fb1fd1960e1bf2507cc4e1cfa9ebe70b`。D02A正本自身で確認した値で、D01Aからの転記ではない。入力byte SHAは末尾記載のcanonical D02Aと一致する。実設定TARGET1000/CYCLES8/GAP.02/ASSOCIATED、DIRECT_DLL1、ADAPT_PRECISION_FROM_START0、ADAPT_C0_WEDGE_DEPTH0、HSD_SOLVE_SECONDS300。

原ログは62427データ行、solve_start/end各208件。lower99回/3951.19秒、upper109回/2057.12秒の完了求解を記録する。これは記録のあるsolve wall合計で、run_endのない強制終了実行の完全な段階時間ではない。

|段階|開始秒|要素数|q|下界call/秒|上界call/秒|
|---|---:|---:|---:|---:|---:|
|coarse_0|1.39|360|4|12 / 172.82|12 / 119.26|
|adapt_1_band|294.89|360|4|10 / 126.16|10 / 106.66|
|band_refine|531.64|454|4|10 / 189.85|10 / 129.23|
|final|852.05|454|8|13 / 158.31|19 / 105.96|
|accuracy_1|1120.75|594|8|14 / 217.95|15 / 102.97|
|accuracy_2|1445.65|886|8|14 / 390.82|12 / 95.02|
|accuracy_3|1936.15|1282|8|13 / 973.25|15 / 378.42|
|accuracy_4|3296.69|1846|8|13 / 1722.04|16 / 1019.59|

Direct144回の内訳は、upper primal26・原dual helper44・fallback7、lower primal13・原dual34・fallback20。Direct不採用時のHSD回帰などを含むので、Direct144回を別の追加Fs208回と合算しない。

**下界が約1.5から増えないこと自体は精度不良ではない。** Python独立対[1.4999745740800001,1.500142126501501]が1.5を挟み、VBAの途中下界1.49999975000022もこの区間内にある。課題は下界を1.5より高くすることではなく、弱い上界約1.681を改善し、既に良い下界をより大きいmeshで繰り返し高コスト探索する制御を避けることにある。最後の3つの下界求解は317.09秒NUMERICAL_FAILURE、278.50秒UNBOUNDED、254.94秒OPTIMALで、Fs1.5付近を反復していた。

worker開始から6060.28秒でRPX_Stop送信、その600.36秒後（累計6660.64秒）に所有Excel PID3880を強制終了しworker exit1。最後の記録は解析開始6052.925531秒の下界solve_end、以降run_endがない。`results_D02A.json`のRUNNINGは更新されず残った状態で、正常完了の証拠ではない。取消不能になった区間がDLL内か組立かは確定できない。

詳細は [D02_LOG_STATS.json](D02_LOG_STATS.json)、[D02_PHASE_TIMINGS.csv](D02_PHASE_TIMINGS.csv)。解析は `python -X utf8 validation/G1R19_D02_route_audit_20261016/analyze_timeout_log.py` で再現できる。

## D02AのPython結果

初期状態から新規計算し、過去のFs/場/根は読み込んでいない。

|項目|結果|
|---|---:|
|全体時間|9.005秒（VBA計算時間ではない）|
|下限|1.4999745740800001 / 304要素|
|上限|1.500142126501501 / 8要素|
|対称GAP|0.011169726928843663%|
|下限の自己根区間|[1.4999745740800001, 1.50002176]|
|上限の自己根区間|[1.5000140015015009, 1.500142126501501]|
|別プロセスの原q8再監査|上下とも物理場・証拠・根が合格|

根の原契約は `width <= 1e-4 * max(1, midpoint)`。
Fsが1.5程度の上限根幅0.000128125はこの契約内であり、絶対幅1e-4と混同しない。

証拠: `validation/G1R19_python26_20261016/results/cold_auto_c0_20261016_2003/`
の`D02A_independent_audit.json`、`cases/D02A/summary.json`、全trialと保存場。
入力byte SHA256は`97c4f88a81575120452a427fdc80621a630a0abf4f09d5da119b3a64f6512a5b`。

## 残る確認

- RPX_Stop送信後にログが途絶えた正確な命令位置・取消確認の欠落区間を追加診断する（現ログだけでは特定不可）。
- 修正VBAでD02Aを新規実行し、両自己根と元q8物理場を検証する。
- VBAから保存した同一メッシュの場をPythonでも再組立・再監査し、制御経路の違いと物理式の相違を区別する。
- 全26新規Python検証と他ケースの回帰を続ける。D02A単独の成功を全26完了へ置き換えない。
