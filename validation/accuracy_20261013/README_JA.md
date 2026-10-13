# 精度改善の数値実験（2026-10-13）

G1R17とIssue #8を基盤に、P06A/P06Nと最新26ケース表の未達ケースを
同じ物理モデルで計算する。実験結果は計算中の候補であり、全ケースの1%達成を
意味しない。VBA候補の実装と実Excel検証は [G1R18](../../versions/G1R18/ADAPT_CONTINUATION.md)。

## 判定

各上下端点は、元の材料・荷重・境界・界面上の応力または速度場を独立に再監査する。
原式の釣合い・降伏・境界残差の許容値1e-7、下界荷重倍率1.000001以上、上界仕事比
0.999999以下を維持する。丸め誤差を包む厳密区間演算の証明ではない。
P06Nを含む非関連入力は既存Davis等価関連モデルであり、直接非関連モデルとの誤差は別。

資料の相対幅 `(U-L)/U` と、VBAの対称相対幅 `2(U-L)/(U+L)` を両方保存する。
本CLIの目標達成には対称相対幅1%以下、q8、各端点の元メッシュでの根幅
`1e-4*max(1, bracket_midpoint)` 以下を要求する。

同じ物理モデルなら上下界のメッシュは異なってよい。新しい試行が失敗しても
既存の最大下界・最小上界をその元メッシュとともに保持する。最適化の停止状態や
監査不合格の値から連続体端点を作らない。原プログラムの有限・許容内の双対場は
離散根の反対符号を案内するだけで、連続体の証拠場にはしない。

## 実装した候補

- 既存のひずみ速度・速度ジャンプによる適合細分化を継続。
- 同一点での回復設定を比較し、物理監査済みの場を優先して保持。
- 各SOCの全行に同じ正倍率を掛ける正規化。双対を元の倍率へ戻して監査。
- SOC正規化の倍率をCSRの行平方和から一括計算。従来のFrobenius倍率と実際の
  下界・上界行列で照合し、元の物理監査を維持した。
- 元の釣合い行列への残差補正後、降伏を含む原物理条件を再監査。補正だけで合格にしない。
- 下界の平衡補正では、組立時の元の行倍率を保持し、原物理式の倍率で100/500/1000回の
  LSMR補正を試す。全原監査に通らなければ、変更していない元の場から従来の1万回補正へ戻る。
  実P06A保存場の補正成功と降伏不合格の棄却、上界の従来補正との完全一致を確認した。
- 同じ合否順位の求解場は、物理監査に通った場の目的値がよいものを保持する。
  不合格の目的値やこの選択から、証拠場や根の符号を生成しない。
- 元目的の双対が許容内で逆符号を確定した点では、上下界証拠にならない主場の
  長時間の補正を省略する。固定荷重・変更した上界目的の双対はこの判定へ使わない。
- 下界の固定荷重倍率による実行可能場の回復。成功場だけを端点に使い、その双対を
  荷重最大化問題の逆符号へ転用しない。
- `upper_proximal_candidate.py`は、保存ベクトルからの変更を小さくする二次目的と、
  元の散逸目的の上限0.999998を組み合わせて上界の物理場を回復する。
  原E/b/G/h/目的/配置のSHA一致と全原q8監査を要求する。補助目的の双対は専用の
  `modified_dual`に保存し、原双対や根符号へ使用しない。case・mesh・model・q・Fsを
  照合し、入力ファイルと実行コードを計算前に固定する。旧Fsの場を中心として使う
  場合も、新Fsでの適合や根完了を移送しない。成功候補の根は未完了のまま保存する。
- c=0自由頂点の角度分割を下界側に限定して試す `--fan-lower`。
- 単一のc=0材料で六頂点の斜面を構成する場合、元の多角形・外周条件・重力を保つ
  薄いすべり領域を上界メッシュへ追加する `c0_slope_wedge_driver.py`。
  解析的なしきい値は探索初期値に限定し、原q8の場の監査と両符号の根を別に確認する。
- 同じモデルのnative26保存場を独立再監査し、履歴の最良場と組み合わせる。
  nativeの成功表示や根完了フラグは引き継がない。旧qの上界も現在のq8散逸で再監査する。
- 目的関数が厳密に0の上界で、元の双対を正倍率で再監査した不可能性rayを
  離散根の案内だけに使う。主場・上下界の合否は変更しない。
- 既に元の原式で符号が確定した点では回復求解を終える。再監査済みの保存場を
  同じメッシュ・同じFsの根符号として使い、同じ点の再求解を省く。
- 不明な点では、元の双対制約を直接解く回復も試す。元のstationarity・cone・目的値を
  再監査できた場合に限って離散根の逆符号に使い、主場の上下界を作らない。
- 摩擦角0の上界に、非圧縮条件と保守的な散逸エピグラフを追加。
  φ>0の旧行列・監査とのビット一致と、純せん断の解析解を確認。
- φ=ψ=0かつ固定荷重0の上界では不変の行列を保持し、Fs依存の価格だけ更新。
  強度比例による場の移送も別に実験し、新Fsでの再監査を必須にした。
- `upper_geometry_driver.py`で同材内の流れに合わせた辺入れ替えを比較する。
  外部・材料境界、座標、材別面積と荷重を保ち、旧証拠は元メッシュへ保存する。
  非入れ子のP2補間は局所座標と内部側の親選択で評価する。補間だけでは採用せず、
  全原監査を通った場だけを初期証拠にできる。通常の`core.prolong`は変更していない。

求解設定は [Clarabelの公式設定](https://clarabel.org/stable/api_settings/)を使用する。
求解器の主方式を置換するVBA/DLL実装ではなく、Pythonでの比較・回復実験である。

## 再実行

リポジトリのルートから、Python3.12/3.13と上記requirementsを使う。Windowsでは
`python -X utf8` を指定する。全ケース入力と初期メッシュはリポジトリ内のG1R17資料を
読み、外部Excelブックなしで実験できる。浸透モデルを乾燥モデルへ置き換える操作は拒否する。

```text
python -X utf8 -m pip install -r validation/accuracy_20261013/requirements.txt
python -X utf8 validation/accuracy_20261013/python/accuracy_driver.py --case P06N --seed validation/accuracy_20261013/seeds/P06N --output validation/accuracy_20261013/results --cycles 8 --max-elements 12000 --fan-lower
python -X utf8 validation/accuracy_20261013/python/accuracy_driver.py --case P06N --resume --recheck --output validation/accuracy_20261013/results --cycles 0
python -X utf8 validation/accuracy_20261013/python/run_suite.py --output validation/accuracy_20261013/suite_results --cycles 8 --max-elements 12000
python -X utf8 validation/accuracy_20261013/python/import_native.py --archive downloads/RPFEM_G1R17_Native26_20261012.zip --output validation/accuracy_20261013/native_seeds --cases P06A P06N P01A P01N P03A P03N P05A P05N P07N D01A D02A D03N D07A P08A P09A P10A
python -X utf8 validation/accuracy_20261013/python/combine_seeds.py --cases P06A P06N --sources validation/accuracy_20261013/native_seeds validation/accuracy_20261013/seeds validation/accuracy_20261013/results --output validation/accuracy_20261013/combined_seeds
python -X utf8 validation/accuracy_20261013/python/lower_accuracy_driver.py --case P06A --seed validation/accuracy_20261013/native_seeds/P06A --method uniform --output validation/accuracy_20261013/results/lower_uniform --rounds 1 --max-elements 24000
python -X utf8 validation/accuracy_20261013/python/c0_slope_wedge_driver.py --case D02A --depth .01 --output validation/accuracy_20261013/results/c0_wedge
python -X utf8 validation/accuracy_20261013/python/c0_slope_wedge_driver.py --case D03N --depth .01 --output validation/accuracy_20261013/results/c0_wedge
python -X utf8 validation/accuracy_20261013/python/upper_proximal_candidate.py --case P03A --seed validation/accuracy_20261013/snapshots/20261014_progress6_audit/supporting_sources/P03A/center_seed --output validation/accuracy_20261013/results/p03a_proximal_replay --alpha .01 --seconds 600
python -X utf8 validation/accuracy_20261013/python/report_accuracy.py --sources validation/accuracy_20261013/native_seeds validation/accuracy_20261013/results --output validation/accuracy_20261013/snapshots/new_audit
python -X utf8 -m unittest discover -s validation/accuracy_20261013/tests -v
python -X utf8 -m unittest discover -s validation/accuracy_20261013/python -p test_*.py -v
python -X utf8 validation/accuracy_20261013/python/check_zero_phi.py
```

`seeds`は添付の監査済み最良場をWindowsで再監査したもの。外部の試行履歴ZIPを
再取り込みする場合は `import_history.py --history 展開先 --output 新しい保存先` を使う。
元の添付ファイルと保存済み履歴は変更しない。

各ケースの`summary.json`に端点、メッシュ/モデルID、元の証拠場、原監査、根の判定、
各段階の履歴がある。各根の`*_trials.json`には不合格・未完了も残す。Python/科学計算
ライブラリの時間とVBA/DLLの実測時間を同じものとして扱わない。

再開時はcase、物理モデル、メッシュ、側、q、政策、Fs、根の符号・順序・幅を照合する。
条件が一致しない根情報は未完了へ降格し、元物理監査を通る場だけを再利用する。

`search(..., initial_dual_guides=[...])`では、case・mode・q・policy・Fs・mesh・modelと
元の双対全ベクトルを指定できる。原プログラムを再組立してstat/cone/分離余裕を
再監査した逆符号点だけを初期の根探索に使い、全ベクトルNPZとSHA256を保存する。
この点はphysical/witnessをfalseとして保存し、主場・連続体上下界を生成しない。
双対点の間は二分探索を使う。保存された物理場と同じFsの競合は拒否する。

`--upper-only`は上下界の保存場がある時に、下界側の証拠を保持して上界メッシュだけを
細分化する実験。`--recheck`指定時の初段は両側の元メッシュの根を再確認する。
`run_suite.py --seed-root 保存場親フォルダ --parallel 2 --fan-lower`で各ケースを
その保存場から計算できる。科学計算ライブラリは各プロセス1スレッドに制限する。
`lower_accuracy_driver.py --method stress_fan --proximity-background 0`で降伏に近い
応力場を重くした指標も比較できる。指標は細分化候補を選ぶだけで精度の証明にはしない。
採用点は`best_<side>_certificate.json`とそのmesh/fieldへ保存する。各段階の計算中も
再監査・集約可能だが、根が未完了の点を1%達成とは数えない。

下界CLIの`--initial-span`は正で有限の幅を指定する。省略時は保存Fsの近傍
`0.002*max(1,Fs)`から始め、必要な側だけ符号探索を広げる。
`run_suite.py --skip-seed-recheck`は保存根の同一性検証を省かず、初期の再求解だけを省く。

途中の独立再監査は `snapshots/20261014_progress6_audit/REPORT_JA.md` と
`COMPARISON_26.csv`に保存した。以前未達だった16ケースのうち14ケースで、幅1%と
両側根完了を確認した。P03Aは0.997594%、P03Nは0.999855%。両者の上界は二次目的で
回復した物理場と、別の元の最適化問題の双対全ベクトルから根を確定した。
`supporting_sources`に元の中心・新しい物理場・補助双対・独立した原双対・根の
全証拠を保存した。補助双対は不使用。P06A/P06Nの改善・VBA本番確認は継続中。
P07Nは0.689432%、薄いすべり領域の候補を使ったD02Aは0.006664%、D03Nは0.008988%。
今回のc=0上界候補は、材料界面を持つP06へ適用しない。
`zero_phi_signed_transport.py`とcapped feasibilityのrayは比較用であり、
後者の追加制約を含む双対を元の根符号へ転用しない。
