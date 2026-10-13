P06N 上界場の補間に対する固定回帰資料

source_mesh.npz は 2914 要素、child_mesh.npz は同じ物理モデルの 6398 要素です。
source_field.npz は Fs=0.6952250450157371 の独立監査済み q=8 上界場です。
legacy_global_inverse_field.npz は改修前の非入れ子補間から採った候補場であり、
原 cone 残差 4.035455428857633e-7 のため上界証拠には採用できません。

child 要素4433 の節点 [23,5] では重心方向の微小 probe が親1938の内部にあります。
旧実装は親1509の probe bary=-1.2479e-11 も探索許容幅内として先に採用しました。
その節点にある両親の許容場の速度差7.31e-11が、細い子要素で増幅します。
新実装は重心方向 probe を局所 bary 座標で評価し、最も内部にある候補を選びます。
P2評価は局所速度差と分割の和が1である性質を使います。
Windowsのnp.longdoubleはfloat64と同じ精度のため、この改善は拡張精度には依存しません。

tests/test_precision_interpolation.py は manifest.json のハッシュを検証し、
旧候補が原監査で拒否され、新しい補間が同じ原行列と許容基準で合格することを調べます。
物理モデル、元場、有限要素次数、監査式、合否許容幅は変更しません。
入れ子での precision_prolong_upper は任意の比較用helperであり、通常core.prolongは維持しています。
