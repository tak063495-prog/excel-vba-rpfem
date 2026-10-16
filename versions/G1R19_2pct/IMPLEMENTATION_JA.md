# G1R19 2%候補（公開版とは別）

基盤は G1R19 と main の Issue #11 修正。公開済み G1R19、Issue11、DLL、保存済み証拠は変更していません。

実際のD01A再現ログでは、全c=0のため従来の散逸指標がゼロとなり、`adapt_band moved=0;weight=0` のまま同じ上界を再探索していました。後段ではq8 DLLも動いていますが、上界約0.700・下界約0.577の幅が改善せず、最後に4016要素の下界候補へ進んでいました。原因を単に「DLL未使用」とはしていません。

変更は次の4モジュールです。

- `RPX_Adapt`：従来散逸指標の総量がゼロの場合、既存のひずみ・速度jump指標へ戻します。正の散逸指標、材料、原方程式、場と根の採用条件は変更しません。
- `RPX_Adapt` / `RPX_UI`：`ADAPT_C0_WEDGE_AUTO` を追加。通常は0、「2%・速度優先の設定」で1になります。対応する乾燥・重力のみ・単一c0材料・6頂点の斜面について、既存の8三角形上界候補を大きな細分化より先に試します。深さは斜面高さの0.001倍。解析的なしきい値は探索初期値に限り、元q8の場・両符号の根の検証でのみ採用します。下界は別の元メッシュ上でq8再探索します。不適格・棄却・未達は従来経路へ進みます。
- `RPX_AccuracyNative` / `RPX_Fs`：HSDのUNBOUNDED status自体は根の証明に使いません。Direct LPの有限候補にも同じ原問題の正規化を適用します。元の自由荷重下界問題のb/h全成分が厳密に0である場合だけ、その同Fs原primalのコピーをλ=1.000003へ正倍率で調整し、全原式と物理門を改めて通します。合格コピーだけを既存のphysical-witness経路へ渡し、不合格時は元の場・状態を保持します。補助双対は使いません。signだけを保証する端点では任意倍率の目的値でsecant補間せず二分法を使います。

Issue #11 の荷重phase復元、原物理門、根幅、別メッシュ所有、単一DLLのABIは維持しています。`ADAPT_PRECISION_FROM_START` の既定値は0のままです。「2%・速度優先の設定」は1へ切り替え、合格したcold条件を再現します。q8を初段から使用する対照試験は `--precision-from-start` を明示し、旧q4前段の試験と区別します。

再生成と試験：

```powershell
python versions/G1R19_2pct/tools/current_sources.py --output-root <新しい候補フォルダー>
python versions/G1R19_2pct/tests/test_overlay_sources.py
python versions/G1R19_2pct/tests/run_adaptive_excel.py --case D01A --guards-only
python versions/G1R19_2pct/tests/run_adaptive_excel.py --case D02A --precision-from-start --budget 900
```

Excel試験は新規所有プロセス1本だけを使用します。ユーザーの既存Excelがあれば起動を拒否します。最後の試験は初期入力から計算し、Python場・根・statusは読み込みません。900秒は協調停止の要求時刻であり、実行中の単一分解が戻るまでの時間を含む厳密なwall上限ではありません。入力・実設定・ソースSHA・DLL SHA・PIDと生成時刻・保存場checkpoint・原行列証拠は新規resultsへ記録します。
