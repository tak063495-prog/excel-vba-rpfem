# D08A：水位に適合する共通静水圧荷重の別実験

元の26ケース実行を変更せず、D08Aだけを独立したディレクトリで新規計算する。材料・外形・外周条件・強度低減方針・指定静水圧 `u=9.81*max(3-y,0)` は保持する。旧nativeの「各解析メッシュの節点uをP1補間」という離散化では、水位を横断する要素を細分化すると体積荷重も変わった。本実験は初期三角形をy=3で分割し、指定された同じ連続静水圧を全ての細分化メッシュで正確に表す。

水位線は材料界面でも速度接着条件でもない。材料IDを保持し、内部辺へ新しい境界条件を加えない。D08Aでは露出地表が水位より高いので外部水圧面力はない。

- 初期214要素の水位横断31要素を分割し、162節点・276要素にした。外形・材料・境界を保持する。
- 下側では体積力 `(0,-gamma+9.81)`、上側では `(0,-gamma)`。固定荷重相は0、全自重・水圧勾配をreference_totalで扱う。
- 初期・1104要素の一様細分化・483要素の局所細分化で、水位横断0、水平体積力厳密0、鉛直合力−2571.4、同じ力の一次モーメントを独立検査した。子の体積荷重は親の該当要素と配列完全一致する。
- 指定されたP1節点圧力の独立勾配との最大差は1.25e-14未満。求解許容の変更ではなく、水位に適合した同じ力学式の正確な評価である。
- 新しいphysical_model_id `984af41b3d0fe16abf5f0ecd1f924014c2086e7b01180d89de4b80ee6cb8117d` を使い、旧nativeの非適合水圧近似の場・根を輸入しない。
- 原q8物理監査、原双対監査、正規化Fs根幅 `(hi-lo)/max(1,(hi+lo)/2)<=1e-4`、対称上下界幅2%を維持する。

[無求解モデル検証](MODEL_VERIFICATION.json)、[分割証拠](data/MESH_PARTITION_AUDIT.json)、[入力・geometryのSHA](INPUTS_26.json)。`tests/test_waterline_conform.py` は共有交点・既存水位節点・材料所有・非適合荷重拒否を検査する。

```powershell
python -X utf8 -m unittest discover -s validation/G1R19_python26_water_20261016/tests -v
python -X utf8 validation/G1R19_python26_water_20261016/verify_water_model.py
python -X utf8 validation/G1R19_python26_water_20261016/guard_water.py --output validation/G1R19_python26_water_20261016/results/NEW_DIRECTORY --seconds 600
```

実行中の記録は results/fresh_common_water_D08A/。全worker treeを作成時刻・完全なコマンドで所有照合し、600秒で実子プロセスから停止する。求解終了後の独立監査は新しいプロセスで原q8式を組み立て直し、元mesh・場・model ID・根端点・符号を検査する。求解時間とその後の無求解監査時間を区別する。新しいVBA変更はこの実験には含めていない。
