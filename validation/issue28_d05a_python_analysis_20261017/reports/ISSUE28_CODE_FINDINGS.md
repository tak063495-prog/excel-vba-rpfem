# Issue #28：D05A上限の原dual復元棄却とスケール差の考察

確認日：2026-10-17（日本時間）。Issue #28本文をGitHubから読み、Fix4製品ソース・同梱DLLソース・Python検証コードを静的に追跡した。Excel・新規最適化は実行していない。製品、プリセット、Gitは変更していない。参照した原文の範囲・SHA256は `ISSUE28_SOURCE_ANCHORS.json` に保存する。

## 結論

1. **エラー名と失敗箇所の混同はソースから確認できる。** HSDの状態がOPTIMALでも、上限値がFs試行点の目標以上で、別の原dual復元が不合格なら `HSD_NOT_CONVERGED` を発生させる。このエラーはHSD自身の未収束だけを意味しない。
2. **ただし、物理監査済みの上限速度場があることだけで、このFs試行点の逆符号を採用してよいわけではない。** 上限の最小化問題で目的値2.40457の許容場は、真の最小値が1より大きいことの証明にならない。反対側Fs符号には監査済みの原dual等が必要である。OPTIMALの文字列だけで復元失敗を無視する変更は、原式に基づく根の保証を弱め得る。
3. **報告されたraw_stat=1429.67127988363は、絶対閾値の僅かなスケール超過という規模ではない。** 仮に最大cost係数と一致するとしても、cost相対閾値 `1e-7×1429.67 ≈ 1.43e-4` を約1,000万倍上回る。さらにcone約4.3e-6、dual目的値約0も、それぞれ独立した受入条件を満たさない。相対閾値へ変えるだけで合格になる候補ではない。
4. **native側のoriginalCost使用は、そのままで直ちに変換漏れを意味しない。** 原dual補助問題は元のcost・元等式・元coneを解き、そのhelperのsolution.xを原dualとして検査する設計である。HSDの正規化dualをこの入口へ直接渡している実装ではない。スケールが数値的な難しさを増している可能性と、原costを使うこと自体が誤りだという主張は分ける必要がある。

## 報告値と直接確認した事実

Issue本文は、Fs=1のHSDが7.07秒・56反復・OPTIMAL、目的値2.40457、primal約2.8e-12、dual約5.0e-11、gap約1.0e-6、物理監査合格、その後の原dual復元がraw_stat=1429.67127988363・raw_cone約4.3e-6・目的値約0、primal復元がraw_eq=1/value=0と報告している。今回、実行機の完全な `analysis.tsv`、full primal/dual、問題行列、solver statusの数値コードは取得していない。これらの値は報告値として扱い、本調査でVBA実行を再現した値とはしない。

`raw_stat ≈ ||cost||∞` と `raw_eq=1/value=0` は、ほぼゼロの候補や有効な解ではないray・打切り点を監査した場合と整合する。しかしcone残差が非ゼロなので、ベクトル全体が厳密なゼロだったとは断定できない。同じ残差が3回出ることは、同一の問題を決定的に解いて同じ棄却候補が返る場合とも整合する。状態の使い回し・メモリ破損・ゼロ配列返却を直接証明するものではない。

## ソースから確定できる制御経路

| 段階 | 原文の位置 | 確認した動作・意味 |
|---|---|---|
| 原上限を解く | `RPX_Fs.bas:248–254` | `RPX_Bound - target` をFs残差にする。q8かつACCURACY_DLL_RECOVERY有効時にnativeEnabledとなる |
| 逆側符号を原式で確認 | `RPX_Fs.bas:271–277` | upperかつOPTIMAL/INFEASIBLEかつ残差>=0なら原dual補助問題を必須にし、不合格時に `HSD_NOT_CONVERGED`。合格したdual guideは物理場とは区別し、xxを消す |
| 同じ復元の再試行 | `RPX_Fs.bas:296–306,889–917` | 回復可能な数値失敗では、再び原dual→primal回復の順で試す。上限では両方不合格なら元のエラーを再送出する。報告の「dual2回＋primal1回」と整合する |
| HSDの判定 | `RPX_Solve.bas:1827–1831` | tau正規化した数値primal/dual残差・gapを使い、数値候補作成後に物理監査へ進む。表示される原目的値・gapにはXQScaleを戻す |
| HSDの候補と物理監査 | `RPX_HSDCandidate.bas:104–110,204–211,255–267` | primal・等式dual・cone dualをtauで割り、上限物理監査が合格した候補をcommitする。ここは「原dualのraw stationarity・cone・分離目的値」の独立監査と同じではない |
| 物理上限の監査 | `RPX_Audit.bas:293–298` | 仕事量1、散逸−固定仕事、物理残差と原目的値の一致を確認する。許容速度場であることを確認し、逆符号用dualを証明する処理ではない |
| 別の原dual問題 | `RPX_AccuracyNative.bas:500–535` | 変数は元等式乗数＋元cone乗数。stationarityの右辺は−originalCost。cone制約とdual目的値>=1.000003を別行で課す |
| 原dualの変数の取り方 | `RPX_AccuracyNative.bas:576–581` | helperのsolution.xがcandidate。helper自身のsolution.zは破棄する。QP/補助問題のdualを原dualと取り違える経路には見えない |
| 原dualの最終受入 | `RPX_AccuracyNative.bas:538–557` | raw stationarity<=1e-7、raw cone<=1e-8、upperの原dual目的値>=1.000003を全て要求する。solver statusだけでは採用しない |
| DLLの完了と最適解の区別 | `native/unified/accuracy/src/lib.rs:108–124,150–157` | 有限のsolution.x/zがあればDONEとして返す。DONEは最適性判定ではなく、solver_statusを別記録する。VBAがその後の原式監査で不適格候補を棄却する設計 |
| VBAの結果読取り | `RPX_AccuracyNative.bas:263–284` | state=2、cancelなし、NResult成功、全値有限なら有限候補とする。ここでゼロや原残差を合格にするわけではなく、最終監査は別 |

VBA位置の共通接頭辞は `versions/G1R20_Fix4/src/vba/`。native位置の接頭辞は `versions/G1R19/`。

## cost・等式行・dualのスケールを区別する

Fix4の `RPX_Conic.bas:412–426` は、原costを保存し、`s = max(1, ||cost||∞)` を `XQScale`、正規化cost `cost/s` を `XQ` とする。等式は `431–446` で原係数・原右辺を保存し、各行の2ノルム `r_i = XEqRowNorm(i)` で割った係数・右辺をHSDへ渡す。`RPX_Solve.bas:290–301` の数値dual残差は、これらの正規化係数について計算される。`RPX_Upper.bas:211–215` はFs更新時にも `XQ*XQScale` で原costを更新する。

したがって、commit済みHSD dualを別途**原問題のdual候補**として利用するなら、等式行とcostを両方戻す必要がある。HSDの等式dualをy_H、cone dualをz_Hとすると、このコードの符号規約では候補は

```
y_raw[i] = s * y_H[i] / r_i
z_raw    = s * z_H
```

となる（HSD候補はcommit前にtau除算済み）。その後、元の未縮約・未正規化問題で `cost + E.T*y_raw - G.T*z_raw`、元cone、原dual目的値を独立に監査する必要がある。row normで戻す既存例は `RPX_Conic.bas:861–905` の `OriginalDualAudit` にある。ただし非ゼロcostのupperについてこの検証をそのまま利用する公開入口はなく、lower用入口は `970–999` で `XQScale=1` を要求している。これは今後の実装候補を調べるための式であり、本調査でHSD dualが原門に合格したと確認したわけではない。

一方、`RPX_AccuracyExport` は原cost・原等式をDLLに送り（`RPX_AccuracyNative.bas:168–181`）、原dual補助問題もそれに対応する変数を直接求める。従って、この新規helperのcandidateにHSD由来の `XQScale` を無条件に掛けるのは正しくない。両方の経路を区別しなければならない。

## Python検証コード側の監査は相対化していない

- `python/experiment_core.py:169–218` は原問題を保管したうえでaffine縮約を行う。原costを `XQScale` で割る処理ではなく、`cost@T` と `cost@off` で縮約する。
- `experiment_core.py:222–277` は縮約dualを元のconeと等式へ復元する。必要な等式乗数をLSMRで求めても、その成否だけを証拠にはしない。
- `experiment_core.py:279–296`、`original_dual.py:31–45` は、**原問題での絶対**stationarity<=1e-7、cone<=1e-8、原dual目的値の条件を要求する。VBAと同じ原式の門であり、Pythonが相対閾値でD05Aを合格にしていたという根拠はない。
- `original_dual.py:57–74` は原dual補助問題のsolution.xを使う。任意の等式行スケーリングは `64–67` で問題の行と右辺を同時に変更するが、最後の `audit_original_dual` は原問題を使う。行スケーリングで門を緩めるものではない。
- `accuracy_driver.py:150–163` は、upperの物理witnessを逆側符号にしない。原dualの条件を満たせないと符号UNKNOWNとなる。`experiment_core.py:288–290` も、物理合格に加えてupper値<=0.999999を物理witnessに要求する。

Python位置の共通接頭辞は `validation/G1R19_python26_20261016/python/`。

## 修正案を評価する際の注意と次の検証

1. **最優先はD05A実行のfull問題・候補を保存すること。** Fs/q/メッシュID・originalCost/XQ/XQScale・等式row norm・HSD commit済みx/y/z・native helperのsolver_status/x/z・実行設定を記録し、同一原式で再監査する。現在の公開本文だけでは、native候補がray、打切り点、数値失敗点のどれだったか決められない。
2. **スケール仮説は原問題の係数をPythonで比較する。** 幾何相似前後のcost・等式・coneの係数範囲、zero dualの残差、報告値と最大cost係数の一致を確認する。別担当のPythonデータがこの候補を検証するが、最大costの一致だけで失敗原因や修正成功を断定しない。
3. **HSD原dualの利用可能性を先に調べる。** 上記のrow normとcostを復元したfull dualが、従来の絶対原門と目標分離値を満たすか、求解をせずに検証する。合格した場合にのみ、その同一Fs・同一問題の逆符号として使う修正を検討できる。
4. **回復失敗をケース打切りと区別する。** 不合格candidateを採用せず、その点をUNKNOWNとして、保存済みの有効な物理場・符号・根区間を壊さず別試行へ進められるかを検討する。HSD OPTIMALを無条件に符号にすることとは違う。
5. **エラーの診断性は別に改善できる。** HSD自身の未収束と、原dual回復棄却を区別し、raw3指標・solver status・段階をログへ残す。これだけで精度や成功率が改善したとは主張しない。
6. **ACCURACY_DLL_RECOVERY=0は切分け試験として扱う。** 完走したとしても、その経路の上下物理場・q8・各自己根をPythonで再監査するまでは修正完了や同等の根保証の証拠にはしない。原門の相対化・不合格dual採用を回避策として提案しない。

本調査で確定したのは「原dual回復の失敗がHSD_NOT_CONVERGEDとして表面化する経路」「正規化の違い」「報告candidateが複数条件で大幅に不合格」という点である。nativeの不良候補が作られた数値的原因と、HSD原dualで回復可能かは未確定である。
