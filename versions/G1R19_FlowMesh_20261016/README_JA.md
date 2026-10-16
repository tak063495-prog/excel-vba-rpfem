# FlowMesh 配列候補API（隔離試作）

`validation/G1R19_python26_20261016/python/mesh_geometry.py::flow_metric_flip` の幾何候補作成を、新規の `RPX_FlowMesh.bas` へ移植した。既存VBA・DLL・実ブックには統合していない。

```vb
accepted = RPX_FlowMeshCandidate(xy, tri, materialId, velocity, protectedEdges, candidateTri, info, 5)
```

入力は0始まりの `xy[n,2] Double`、`tri[ne,3] Long`、`materialId[ne] Long`、**物理速度だけの12*ne個**のDouble配列。`protectedEdges` は `"小node:大node"` をキーとする明示的なDictionaryで、既存RConstraintsを全て渡す。追加保護がなければ空Dictionaryを渡す。外周と材料境界はAPI自身でも保護する。出力は入力とは別の未使用配列・Object変数を使い、入力triと出力triのaliasを拒否する。

元meshのP2速度を重心で微分したJ.TJから、固有値をhi/9で下限処理したmetricを一度だけ作る。flip後も元element IDに対応するmetricを固定し、各passの候補をgain降順・同点は辺の発見順に処理する。同じpassで一度使った要素は再利用しない。既定5pass、APIの上限も5pass。

同材の凸な2三角形cavityだけを対象とし、正向き・元面積・保護辺を維持する。metric品質は `after>1.05*before+1e-10`、等方幾何品質は `after>=.6*before`。全材料面積と外周/材料辺の保存を最後にも確認してから候補を返す。2x2固有値下限は同じスペクトル関数を解析式で評価するため、NumPyのeigenvector計算との丸め差が同点付近の選択へ影響する可能性は明示して比較する。

このAPIはglobal model・field・Fs・rootを読まず、書き換えない。`info("certificate")=False`、`info("nested")=False`。旧速度場や旧根を新meshへ採用しない。新meshを実装側で採用する場合は、独立した新規求解・原q8物理監査・同meshのown-root完了が別途必須である。

## Solver0 の比較手順

Pythonのfixture生成は既存のP06上界場を幾何方向としてだけ使う。元場を原q8で検査してから、元Python関数で5pass候補を保存する。Nには内部RConstraintsを想定した追加保護辺も与える（元関数へ辺の拒否predicateだけを追加した参照）。新たな最適化や根探索は行わない。

実Excelはintegration担当が所有PIDを限定した使い捨てブックで行う。このnew moduleとtest-only `tests/FlowMeshArrayProbe.bas` を追加し、compile/save/reopen/compileする。各fixtureについて以下を実行する。

```vb
result = FlowMeshFileProbe(fixtureInputPath, newOutputPath)
```

probeはsource配列が不変であることを確認し、候補tri/flipinfoだけをbinary出力する。meshをglobalへinstallしない。元bookの物理model/Fs/root状態もharness側で呼出前後に比較する。全出力を保管して次のPython比較を行う。

```powershell
D:/Work/rpfem-venv/Scripts/python.exe -X utf8 versions/G1R19_FlowMesh_20261016/tests/compare_excel.py --fixtures versions/G1R19_FlowMesh_20261016/tests/fixtures --outputs <NEW_OUTPUT_FOLDER>
```

Python期待値とのconnectivity完全一致を最初の判定とする。差があれば出力を保持し、丸めの同点・順序・metricのどこが原因かを調べる。幾何が有効でも自動で合格へ読み替えず、新版の参照差分として説明する。現時点ではVBA compile/runtimeは未確認である。
