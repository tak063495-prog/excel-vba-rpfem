# 実Excel micro手順（未実行）

既存のD08、4fix、WaterCommon実行・保存ブックを使わず、所有PIDを限定した新しい使い捨てブックを使う。マクロ実行はintegration担当へ引き渡す。本overlayはまだ実Excel compile/reopenを主張しない。

## 準備と出力の分離

1. 製品候補用62モジュールは `G1R19 → G1R19_2pct → G1R19_StateReset_20261016` の順に同名置換して作る。4モジュール以外の基準を変更しない。新規候補ブックへimportし、compile/save/close/reopen/compileし、全62のsource/loaded/reopened identityを確認する。公開ブックは置き換えない。
2. micro用は以下のgeneratorで**別の新規フォルダ**へ63モジュールを作る。4候補モジュールを元にConic/Solveへprivate状態確認関数を追加し、Compile入口とSeepSolve入口に停止trapを加える。製品overlay内の4モジュールにはtrapもsentinelも入らない。

```powershell
D:/Work/rpfem-venv/Scripts/python.exe -X utf8 versions/G1R19_StateReset_20261016/tests/make_probe_sources.py --output <NEW_TEST_SOURCE_DIRECTORY>
```

3. 63モジュールのtest-onlyコピーを、さらに別の使い捨てブックへimport/compileする。TEST_ONLY_MANIFEST.json、実importしたmodule SHA、owned PID/creationtime、DLL SHA、各呼出しの戻り値とexit確認を保存する。このブックから場・根・数値証拠を採用しない。
4. 比較用negative controlは同じgeneratorの `--negative-control` で作る。これは4行をCall化する前のsourceに同じinstrumentationだけを追加する。旧版の対応probeがFAILし、修正版でPASSすることが望ましい。失敗の理由が意図したsentinel/入口検査であることまで記録し、compile失敗や入力不備を代わりに数えない。

## 入口で実際に処理されたことの検査

| 呼出し | 前提・sentinel | 必須の観測 |
|---|---|---|
| `SRProbeConicEntry()` | 同じConicモジュール内でprivate originalGuideReady=True、model=TEST_ONLY_READY_SENTINELを設定。意図的にscope不適格なcandidateをGateへ渡す。 | GateはFalse、ready=False、model空。通常の候補監査に進む前に実Clearが呼ばれたことを確認。保存した従前値はcleanupで復元する。 |
| `SRProbeDirectEntry("lower",1.3)` / `("upper",1.3)` | fresh dry D02入力・mesh、idle、RR0、RFsActive=True、q8、Direct/Accuracy有効、AF無効。RProgramKind空。Solve内のoriginalK/warm/cache readyフラグを全てTrueにし、対応配列を確保したtest-only sentinelを置く。 | RPX_Boundから実際にDirect新規組立経路へ入る。Compile入口trapが1回、error18 TEST_ONLY_STOP_BEFORE_COMPILE、全readyフラグFalse・全sentinel配列消去・構造世代増分を確認。Compile本体・DLL solver・HSDは0回。各probe前にfresh mesh/input scopeを用意する。 |
| `SRProbeKrylovEntry()` | usable native DLLが存在すること。正常なinput scopeでXCancel=Trueを設定。 | HNativeKrylov入口がerror18 ANALYSIS_CANCELLEDで止まる。NativeKrylovBegin/Stepより前にHContextの取消検査を通る。test-copyでは同じ物理行のHContext直後、HEnsureDDよりも前に別trapを置き、旧版が検査を飛ばしても配列準備・未初期化pointer・native kernelへ進ませない。このtrapのhitは0が必須。実native solver呼出しなし。DLLなしはPASSとしない。 |
| `SRProbeSeepEntry()` | Defaults対象の7値を全て−777へ変更、Compileとは独立のSeepSolve入口trapを有効化。 | SeepTestRelax(.25)から入る。SeepSolve本体より前にdefault6値＋指定relax .25へ置換済み、trap1回/error18 TEST_ONLY_STOP_BEFORE_SEEP_SOLVE。水理FEM求解0回。 |

Direct probe前の標準InputBegin/ReadInputs/GenerateMesh/荷重phase等は既存のmicro harnessと同じ経路で用意する。自己設定された偽のRootValid/PhysicalGateAcceptedを用いて成功を作らない。probeは意図的な取消を返すため、harness側では期待された専用メッセージまで一致させる。

sentinelは失効動作の検査に限る。いずれのPASSも物理場・Fs根・精度GAPの証明ではない。cleanupは作業配列やcurrent flagsを失効させるので、probe後のmesh/fieldを製品実験へ引き継がない。

## その後の数値確認

microだけで製品を採用しない。solverを使う確認が別途許可・割当された後、test-only instrumentationを全て除いた62モジュール候補で、固定したD01/D02の新規coldと原q8物理監査・own rootsを再確認する。既存のD01/D02の成功した保存証拠は元のsource hashの証拠として保持し、新4行版の実行証拠へ付け替えない。
