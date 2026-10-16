# D08・同一DLLと同一原LPの設定比較

対象は実VBA WaterCommon実行 `owned_excel_1792152100656896200` が保存した、上界429要素・Fs=1.3419713973999023 のAC18P001原LP。LPとcontextとDLLを安定した別snapshotへコピーし、同じCSR・原cost・原rhsをそのまま使った。追加制約、行スケール、場の正規化、Fs変更はない。候補の採用や根証明は行わない診断である。

- LP SHA256: `48d81277333724e486ed49726d75ad3cb412cb5bce4296e1c68c0ca4fe607ce8`
- DLL SHA256: `66da1a21be43cd2b3a3bfbaa9d13aad5fb1fd1960e1bf2507cc4e1cfa9ebe70b`
- 変数5,148、行16,324、非零166,119。threads=1、各time_limit=30秒、実時間は各0.63～0.90秒。
- native ExactModelの原doubleから再構成したmesh/body/material/boundaryを使用。正本D08材料・境界、水位3・水単位体積重量9.81の共通水圧荷重を照合し、面積158、合力[0,−2571.4]。乾燥モデルへのfallbackなし。

| equilibrate | 内部tol | 原等式残差 | 原primal錐残差 | 原dual停留残差 | 独立q8物理場 |
|---:|---:|---:|---:|---:|---|
| 0 | 1e-13 | 4.61648e-10 | 7.43109e-8 | 1.89336e-5 | PASS、値0.9984903404806245 |
| 1 | 1e-13 | 4.13411e-12 | 3.73728e-10 | 2.37082e-9 | PASS、値0.9985436606601354 |
| 0 | 1e-10 | 4.61648e-10 | 7.43109e-8 | 1.89336e-5 | PASS、値0.9984903404806245 |
| 1 | 1e-10 | 4.13411e-12 | 3.73728e-10 | 2.37082e-9 | PASS、値0.9985436606601354 |

equilibrate=0/tol=1e-13の独立FFI結果xは、元VBAの保存primal xと全5,148要素のuint64が一致した。したがって、このLPではPython→同DLLとVBA→同DLLに数値結果の差はない。equilibrationを変えると原錐残差が約200分の1、dual停留残差が約8,000分の1になる。内部tolの変更だけではこの点の結果は変わらない。

全設定のsolver status整数は4（metadata）、iterationはequ0が26、equ1が40。statusを証拠として扱っていない。原primalの閾値1e-7、原dual停留1e-7・錐1e-8、独立q8物理監査をそのまま評価した。dualの残差合格だけで必要な根の符号が得られたとは主張しない。

これは既にprimal原門へ合格した保存LPであり、実行ログにあった約7e-7の棄却点そのものの改善再現ではない。equ1が有望な設定候補である一方、採用前には当該棄却Fsの正確な原LPと元FieldGateでの確認が必要である。全26やD08のGAP達成証明にも数えない。active VBA・ブック・source・既存数値出力は変更していない。

結果と全ベクトルは `results/original_upper429_equ_tol_v2_20261016/RESULTS.json` と4個のNPZに保存。CLIの最初のpreflightはD02限定のzero-cost条件を誤って持ち込んだため、D08の非零costを拒否して求解前に終了した。D08は非零costを持つ正当なupper original LPなので、velocity-only/soil-onlyの形状検査と全原costの保持へ直した。最初のsnapshotは `results/original_upper429_equ_tol_20261016` に残し、数値成功として扱っていない。

再現例（出力は毎回新規フォルダを指定する）:

```powershell
D:/Work/rpfem-venv/Scripts/python.exe -X utf8 validation/G1R19_D08_ffi_20261016/probe_original_lp.py --lp validation/G1R19_WaterCommon_excel_20261016/results/owned_excel_1792152100656896200/RPFEM_logs/20261016_210153_9228_analysis.tsv.accuracy_primal_solve_8_34_upper.bin --dll validation/G1R19_WaterCommon_excel_20261016/results/owned_excel_1792152100656896200/rpfem_native_x64.dll --output <NEW_DIRECTORY> --seconds 30
```
