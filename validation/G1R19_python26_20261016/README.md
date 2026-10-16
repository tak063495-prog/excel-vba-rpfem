# G1R19 26ケース・2%検証（2026-10-16）

対象はG1R17 native26と同じ26モデル。G1R19の元の物理式・原q8監査を用い、
Pythonの適応細分化・原根探索で対称上下界幅2%以下へ到達するか、新規計算で確認する。
過去の保存場の再監査と、初期状態からの新規計算は別表に記録し、合算しない。

`2*(U-L)/(U+L) <= .02` に加え、両側の原物理場・q8・各own-mesh根区間について
`width <= 1e-4 * max(1, midpoint)` を要求する（既存の正規化された原契約）。
入力manifestのabsoluteというラベルの訂正は `ROOT_CONTRACT_ERRATA.md` に記録した。
原物理残差/raw下界残差1e-7、原双対のstationarity1e-7/cone1e-8は変更しない。
NケースはDavis等価関連モデル、Aケースは強度低減時の関連方針。
D08Aの固定水圧も保持し、乾燥化による代用を禁止する。

実行は `run_26.py --output NEW_DIRECTORY`。既定は2プロセス・各solver1 thread、
1ケース900秒・1求解30秒、根幅1e-4、最大8改善回・32,000要素。
in-flight assembly/LSMRを含むハード上限は900+60秒。
Windows venvの親redirectorだけでは実Pythonが残るため、今回の実行では
`guard_batch.py` を独立に併走し、900+50秒で所有する実Python子→親の順に停止する。
所有PID・作成時刻・完全なコマンドを再照合し、保存済み証拠のSHAを記録する。
時間切れ・不完全根・不正場・未実行を合格にしない。各ケースの終了後、別プロセスで元q8式を組み立て直して再監査する。

上下界を小数点以下2桁へ四捨五入した値と、その一致も別列に記録する（Decimal ROUND_HALF_UP）。
2%判定と丸め一致は別の指標で、丸め一致を追加の終了目標にはしない。

これはPython参照解法の到達性検証であり、VBAのLEGACY制御経路を同じ順序で実行するものではない。
G1R19のVBA側で見つかった差分と修正後の実Excel結果は、別途記録する。

初期調査で確認した報告: [PR #19・D01A 90分TIMEOUT](https://github.com/tak063495-prog/excel-vba-rpfem/pull/19)。
同じ公式ブック・単一DLL・ASSOCIATED・GAP=.02で、GAP約19.22%のまま208求解に達して停止している。
原ログは `validation/G1R19_preset2pct_20261016/D01A/` に保全されている。

進行中の26ケース表は `CURRENT_RESULTS.csv`、全試行履歴は `CURRENT_RESULTS.json`。
`collect_current.py --batch results/cold_auto_c0_20261016_2003` で更新する。
別版で新規再計算したケースは `COLD_RETRIES.json` に明示して追加し、最初の失敗行も履歴に残す。
集計時には独立監査が参照した全入力・mesh・場のSHAも再照合する。
過去の16保存済みペアの再監査は `EXISTING_SAVED_16_REAUDIT.json` に別記し、26新規計算には数えない。

原因調査と確認済みの修正結果は [CURRENT_FINDINGS_JA.md](CURRENT_FINDINGS_JA.md)。
別のsupervisorが管理する新規実験は `--supervised-batch PATH` で集計できる。
supervisor・solver・独立監査の正常終了、sourceの開始/終了/現在SHA一致を確認し、
途中checkpointの幅が2%でも正常完了には数えない。

表は `scientific_pair_verified`（保存対を別プロセスで元式・両自己根まで確認）と
`cold_execution_verified`（初期状態からのdriverも正常終了）を区別する。
前者だけがtrueの場合、時間切れ・worker終了コード未記録などの理由をstatusに残す。
この場合はGAP到達の証拠にはなるが、規定時間内の正常完了や本番での速度を保証しない。

既知メッシュの形状だけからの新規求解は `--geometry-pair FOLDER` で明示追加する。
この分類では過去応力・速度・Fsを入力せず、実際の各componentの開始/終了ソース・入力・場・根の記録を照合する。
後段で集めた対を単一のcold実行と表示することはない。過去保存場そのものの再生も26例の新規求解には数えない。

P06Aでは既知形状から新しい場・根を求め、GAP1.8173754281%を別プロセス原q8監査で確認した。
`composite_provenance.py` は後付けで架空の開始manifestを作らず、各求解componentの実記録へ結び付ける。
従来のPARTIAL_24of26 ZIPはその時点の不変snapshotとして保持する。
