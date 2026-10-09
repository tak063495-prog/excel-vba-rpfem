# SRM高速化のテストケース集 — 26入力

作成日: 2026-10-01。文献記号R01～R11は主計画/Source_Audit.md/sources.jsonと共通。

## 共通規則

単位はm、kN、kPa、角度degree、単位奥行1m。平面ひずみ・小ひずみ・弾完全塑性MC。底面固定、左右の垂直端は水平変位拘束、その他の露出面は自由。P07は右垂直端を持たない。全外周辺をJSONで明示し、固定条件を推測で付けない。

**Nは非関連、Aは関連。Aでは常にpsi_F=phi_F。Fs<1を許可する。論文参照値は方法別観測であり、当該有限要素メッシュの保証値ではない。** 同じ物理条件の基準解は未実施なのでnull/NOT_RUNとした。

追加tension cutoffは明示的になし。これはMC自体の許容応力域を無視するという意味ではなく、別の引張面を足していないという意味。論文が引張オプションを明示しない場合はこの設定を補足仮定として扱う。

M0/M1/M2は粗/中/細の目標h。現行要素・積分をmanifestから継承し、Python/VBAで同じメッシュを使用する。未対応機能はSKIP_WITH_REASON。自動的にψ=φ、DP、微小c、別積分へ変換しない。

## 一覧

| ID | ケース | 流れ則 | 文献/由来 |
|---|---|---|---|
| P01A | Jia2024 / Cheng homogeneous Case 1, psi=0, associated twin | associated | R02 / derived_associated_twin |
| P01N | Jia2024 / Cheng homogeneous Case 1, psi=0 | nonassociated | R02 / literature_geometry_parameters |
| P02A | Jia2024 / Cheng homogeneous Case 2, psi=0, associated twin | associated | R02 / derived_associated_twin |
| P02N | Jia2024 / Cheng homogeneous Case 2, psi=0 | nonassociated | R02 / literature_geometry_parameters |
| P03A | Jia2024 / Cheng homogeneous Case 3, psi=0, associated twin | associated | R02 / derived_associated_twin |
| P03N | Jia2024 / Cheng homogeneous Case 3, psi=0 | nonassociated | R02 / literature_geometry_parameters |
| P04A | Jia2024 / Cheng homogeneous Case 4, psi=0, associated twin | associated | R02 / derived_associated_twin |
| P04N | Jia2024 / Cheng homogeneous Case 4, psi=0 | nonassociated | R02 / literature_geometry_parameters |
| P05A | Three horizontal layers | associated | R02 / derived_associated_twin |
| P05N | Three horizontal layers | nonassociated | R02 / literature_with_explicit_psi_policy |
| P06A | Thin frictional weak band, 28 m domain | associated | R02, R10, R11 / cross_checked_geometry_with_unreported_tensile_policy |
| P06N | Thin frictional weak band, 28 m domain | nonassociated | R02, R10, R11 / cross_checked_geometry_with_unreported_tensile_policy |
| P07A | Griffiths–Lane Example1, D=1 | associated | R03 / nondimensional_literature_case_dimensionalized |
| P07N | Griffiths–Lane Example1, D=1 | nonassociated | R03 / nondimensional_literature_case_dimensionalized |
| P08A | Uniform undrained clay | associated | R03 / nondimensional_literature_case_dimensionalized |
| P09A | Weak clay foundation | associated | R03 / nondimensional_literature_case_dimensionalized |
| P10A | Clay foundation mechanism transition | associated | R03 / nondimensional_literature_case_dimensionalized |
| D01A | All c=0, 45 degree face, phi=30 | associated | R02 / designed_diagnostic |
| D02A | All c=0, slope ratio 2:3, phi=45 | associated | 独自診断 / designed_diagnostic |
| D03N | D02 geometry, c=0 nonassociated psi=0 | nonassociated | 独自診断 / designed_diagnostic |
| D04A | Elastic modulus sensitivity E/10 | associated | R02 / designed_diagnostic |
| D05A | Geometric similitude: L x1000, gamma /1000 | associated | R02 / designed_diagnostic |
| D06A | Translation covariance (+1000,+2000) | associated | R02 / designed_diagnostic |
| D07A | Crest surcharge 20 kPa on x=2..8 | associated | R02 / designed_diagnostic |
| D08A | Fixed pore pressure field, water elevation y=3 | associated | R02 / designed_diagnostic |
| D09A | Near-incompressibility diagnostic nu=0.49 | associated | R02 / designed_diagnostic |

## 使用順序

**最小smoke群:** P01N, P02N, P04N, P07N, P08A, D01A, D02A。まずM0で入力/収束経路を確認する。
**保留評価群:** P03N, P05N, P06N, P09A, P10A。これらに合わせて探索パラメータを再調整しない。
**上界の適格性:** P01A～P04A、D01A/D02Aで関連機構を検査。P05/P06は単一くさび適用を拒否。Nは関連代理モデルのHINTにとどめる。
**異常系/不変性:** D03N～D09A。対応機能がないD08A等は未対応と報告し、検査済み扱いにしない。

## 数値参照の読み方

R02 Table1の4列は別の土質定数を用いたCase1～4。3.1～3.4節の4種類の地形例とは別である。A派生にN参照値を移していない。R03のTable2は反復挙動の比較資料であり、現在のソルバーの反復数目標ではない。

薄層P06の下側境界は(23,4.5)から(28,5)まで続く。R10の反転図で照合した。引張政策等の相違が残るので、eSPFEM0.90だけを厳密正解にはしない。[R02,R10]

## 各入力の詳細

### P01A — Jia2024 / Cheng homogeneous Case 1, psi=0, associated twin

- 出典箇所: p8 Fig3 Table1
- 由来区分: `derived_associated_twin`。面積: **158 m²**。
- 外形（反時計回り）: `(0, 0) → (20, 0) → (20, 4) → (16, 4) → (10, 10) → (0, 10)`
- psi政策: `equal_reduced_phi`。各Fs独立自重試行を文献比較用の基本とする。
- 実解析/VBA: **NOT_RUN**。

| 材料 | E(kPa) | ν | γ(kN/m³) | c(kPa) | φ(°) | ψ初期(°) |
|---|---:|---:|---:|---:|---:|---:|
| soil | 14000 | 0.3 | 20 | 10 | 5 | 5 |

領域:
- `soil`: `(0, 0) → (20, 0) → (20, 4) → (16, 4) → (10, 10) → (0, 10)`

メッシュ目標h: M0=1.25m, M1=0.625m, M2=0.3125m。

この派生入力/診断入力の定量Fsは未設定。元ケースから参照値をコピーしていない。

補足仮定・制限:
- R02 uses explicit large-deformation SPFEM; compare failure onset, not its 6 s final displacement.
- No explicit tensile cap was specified in the inspected benchmark description; this input adds none.
- Fs<1 must be allowed; the independent-gravity trial profile is an explicit static-FEM adaptation.

### P01N — Jia2024 / Cheng homogeneous Case 1, psi=0

- 出典箇所: p8 Fig3 Table1
- 由来区分: `literature_geometry_parameters`。面積: **158 m²**。
- 外形（反時計回り）: `(0, 0) → (20, 0) → (20, 4) → (16, 4) → (10, 10) → (0, 10)`
- psi政策: `hold_initial`。各Fs独立自重試行を文献比較用の基本とする。
- 実解析/VBA: **NOT_RUN**。

| 材料 | E(kPa) | ν | γ(kN/m³) | c(kPa) | φ(°) | ψ初期(°) |
|---|---:|---:|---:|---:|---:|---:|
| soil | 14000 | 0.3 | 20 | 10 | 5 | 0 |

領域:
- `soil`: `(0, 0) → (20, 0) → (20, 4) → (16, 4) → (10, 10) → (0, 10)`

メッシュ目標h: M0=1.25m, M1=0.625m, M2=0.3125m。

| 報告方法 | Fs | 性格 |
|---|---:|---|
| Cheng2007_LEM | 0.65 | R02 Table 1; secondary when other authors are named |
| Cheng2007_FEM | 0.69 | R02 Table 1; secondary when other authors are named |
| Zhou2022_LB_as_reported | 0.68 | R02 Table 1; secondary when other authors are named |
| Zhou2022_UB_as_reported | 0.7 | R02 Table 1; secondary when other authors are named |
| Zhou2022_SOCP | 0.69 | R02 Table 1; secondary when other authors are named |
| Wang2021_SOCP | 0.67 | R02 Table 1; secondary when other authors are named |
| Jia2024_eSPFEM | 0.68 | R02 Table 1; secondary when other authors are named |

補足仮定・制限:
- R02 uses explicit large-deformation SPFEM; compare failure onset, not its 6 s final displacement.
- No explicit tensile cap was specified in the inspected benchmark description; this input adds none.
- Fs<1 must be allowed; the independent-gravity trial profile is an explicit static-FEM adaptation.

### P02A — Jia2024 / Cheng homogeneous Case 2, psi=0, associated twin

- 出典箇所: p8 Fig3 Table1
- 由来区分: `derived_associated_twin`。面積: **158 m²**。
- 外形（反時計回り）: `(0, 0) → (20, 0) → (20, 4) → (16, 4) → (10, 10) → (0, 10)`
- psi政策: `equal_reduced_phi`。各Fs独立自重試行を文献比較用の基本とする。
- 実解析/VBA: **NOT_RUN**。

| 材料 | E(kPa) | ν | γ(kN/m³) | c(kPa) | φ(°) | ψ初期(°) |
|---|---:|---:|---:|---:|---:|---:|
| soil | 14000 | 0.3 | 20 | 10 | 25 | 25 |

領域:
- `soil`: `(0, 0) → (20, 0) → (20, 4) → (16, 4) → (10, 10) → (0, 10)`

メッシュ目標h: M0=1.25m, M1=0.625m, M2=0.3125m。

この派生入力/診断入力の定量Fsは未設定。元ケースから参照値をコピーしていない。

補足仮定・制限:
- R02 uses explicit large-deformation SPFEM; compare failure onset, not its 6 s final displacement.
- No explicit tensile cap was specified in the inspected benchmark description; this input adds none.
- Fs<1 must be allowed; the independent-gravity trial profile is an explicit static-FEM adaptation.

### P02N — Jia2024 / Cheng homogeneous Case 2, psi=0

- 出典箇所: p8 Fig3 Table1
- 由来区分: `literature_geometry_parameters`。面積: **158 m²**。
- 外形（反時計回り）: `(0, 0) → (20, 0) → (20, 4) → (16, 4) → (10, 10) → (0, 10)`
- psi政策: `hold_initial`。各Fs独立自重試行を文献比較用の基本とする。
- 実解析/VBA: **NOT_RUN**。

| 材料 | E(kPa) | ν | γ(kN/m³) | c(kPa) | φ(°) | ψ初期(°) |
|---|---:|---:|---:|---:|---:|---:|
| soil | 14000 | 0.3 | 20 | 10 | 25 | 0 |

領域:
- `soil`: `(0, 0) → (20, 0) → (20, 4) → (16, 4) → (10, 10) → (0, 10)`

メッシュ目標h: M0=1.25m, M1=0.625m, M2=0.3125m。

| 報告方法 | Fs | 性格 |
|---|---:|---|
| Cheng2007_LEM | 1.3 | R02 Table 1; secondary when other authors are named |
| Cheng2007_FEM | 1.36 | R02 Table 1; secondary when other authors are named |
| Zhou2022_LB_as_reported | 1.31 | R02 Table 1; secondary when other authors are named |
| Zhou2022_UB_as_reported | 1.36 | R02 Table 1; secondary when other authors are named |
| Zhou2022_SOCP | 1.33 | R02 Table 1; secondary when other authors are named |
| Wang2021_SOCP | 1.32 | R02 Table 1; secondary when other authors are named |
| Jia2024_eSPFEM | 1.3 | R02 Table 1; secondary when other authors are named |

補足仮定・制限:
- R02 uses explicit large-deformation SPFEM; compare failure onset, not its 6 s final displacement.
- No explicit tensile cap was specified in the inspected benchmark description; this input adds none.
- Fs<1 must be allowed; the independent-gravity trial profile is an explicit static-FEM adaptation.

### P03A — Jia2024 / Cheng homogeneous Case 3, psi=0, associated twin

- 出典箇所: p8 Fig3 Table1
- 由来区分: `derived_associated_twin`。面積: **158 m²**。
- 外形（反時計回り）: `(0, 0) → (20, 0) → (20, 4) → (16, 4) → (10, 10) → (0, 10)`
- psi政策: `equal_reduced_phi`。各Fs独立自重試行を文献比較用の基本とする。
- 実解析/VBA: **NOT_RUN**。

| 材料 | E(kPa) | ν | γ(kN/m³) | c(kPa) | φ(°) | ψ初期(°) |
|---|---:|---:|---:|---:|---:|---:|
| soil | 14000 | 0.3 | 20 | 20 | 5 | 5 |

領域:
- `soil`: `(0, 0) → (20, 0) → (20, 4) → (16, 4) → (10, 10) → (0, 10)`

メッシュ目標h: M0=1.25m, M1=0.625m, M2=0.3125m。

この派生入力/診断入力の定量Fsは未設定。元ケースから参照値をコピーしていない。

補足仮定・制限:
- R02 uses explicit large-deformation SPFEM; compare failure onset, not its 6 s final displacement.
- No explicit tensile cap was specified in the inspected benchmark description; this input adds none.
- Fs<1 must be allowed; the independent-gravity trial profile is an explicit static-FEM adaptation.

### P03N — Jia2024 / Cheng homogeneous Case 3, psi=0

- 出典箇所: p8 Fig3 Table1
- 由来区分: `literature_geometry_parameters`。面積: **158 m²**。
- 外形（反時計回り）: `(0, 0) → (20, 0) → (20, 4) → (16, 4) → (10, 10) → (0, 10)`
- psi政策: `hold_initial`。各Fs独立自重試行を文献比較用の基本とする。
- 実解析/VBA: **NOT_RUN**。

| 材料 | E(kPa) | ν | γ(kN/m³) | c(kPa) | φ(°) | ψ初期(°) |
|---|---:|---:|---:|---:|---:|---:|
| soil | 14000 | 0.3 | 20 | 20 | 5 | 0 |

領域:
- `soil`: `(0, 0) → (20, 0) → (20, 4) → (16, 4) → (10, 10) → (0, 10)`

メッシュ目標h: M0=1.25m, M1=0.625m, M2=0.3125m。

| 報告方法 | Fs | 性格 |
|---|---:|---|
| Cheng2007_LEM | 1.06 | R02 Table 1; secondary when other authors are named |
| Cheng2007_FEM | 1.2 | R02 Table 1; secondary when other authors are named |
| Zhou2022_LB_as_reported | 1.18 | R02 Table 1; secondary when other authors are named |
| Zhou2022_UB_as_reported | 1.2 | R02 Table 1; secondary when other authors are named |
| Zhou2022_SOCP | 1.19 | R02 Table 1; secondary when other authors are named |
| Wang2021_SOCP | 1.16 | R02 Table 1; secondary when other authors are named |
| Jia2024_eSPFEM | 1.18 | R02 Table 1; secondary when other authors are named |

補足仮定・制限:
- R02 uses explicit large-deformation SPFEM; compare failure onset, not its 6 s final displacement.
- No explicit tensile cap was specified in the inspected benchmark description; this input adds none.
- Fs<1 must be allowed; the independent-gravity trial profile is an explicit static-FEM adaptation.

### P04A — Jia2024 / Cheng homogeneous Case 4, psi=0, associated twin

- 出典箇所: p8 Fig3 Table1
- 由来区分: `derived_associated_twin`。面積: **158 m²**。
- 外形（反時計回り）: `(0, 0) → (20, 0) → (20, 4) → (16, 4) → (10, 10) → (0, 10)`
- psi政策: `equal_reduced_phi`。各Fs独立自重試行を文献比較用の基本とする。
- 実解析/VBA: **NOT_RUN**。

| 材料 | E(kPa) | ν | γ(kN/m³) | c(kPa) | φ(°) | ψ初期(°) |
|---|---:|---:|---:|---:|---:|---:|
| soil | 14000 | 0.3 | 20 | 20 | 35 | 35 |

領域:
- `soil`: `(0, 0) → (20, 0) → (20, 4) → (16, 4) → (10, 10) → (0, 10)`

メッシュ目標h: M0=1.25m, M1=0.625m, M2=0.3125m。

この派生入力/診断入力の定量Fsは未設定。元ケースから参照値をコピーしていない。

補足仮定・制限:
- R02 uses explicit large-deformation SPFEM; compare failure onset, not its 6 s final displacement.
- No explicit tensile cap was specified in the inspected benchmark description; this input adds none.
- Fs<1 must be allowed; the independent-gravity trial profile is an explicit static-FEM adaptation.

### P04N — Jia2024 / Cheng homogeneous Case 4, psi=0

- 出典箇所: p8 Fig3 Table1
- 由来区分: `literature_geometry_parameters`。面積: **158 m²**。
- 外形（反時計回り）: `(0, 0) → (20, 0) → (20, 4) → (16, 4) → (10, 10) → (0, 10)`
- psi政策: `hold_initial`。各Fs独立自重試行を文献比較用の基本とする。
- 実解析/VBA: **NOT_RUN**。

| 材料 | E(kPa) | ν | γ(kN/m³) | c(kPa) | φ(°) | ψ初期(°) |
|---|---:|---:|---:|---:|---:|---:|
| soil | 14000 | 0.3 | 20 | 20 | 35 | 0 |

領域:
- `soil`: `(0, 0) → (20, 0) → (20, 4) → (16, 4) → (10, 10) → (0, 10)`

メッシュ目標h: M0=1.25m, M1=0.625m, M2=0.3125m。

| 報告方法 | Fs | 性格 |
|---|---:|---|
| Cheng2007_LEM | 2.24 | R02 Table 1; secondary when other authors are named |
| Cheng2007_FEM | 2.28 | R02 Table 1; secondary when other authors are named |
| Zhou2022_LB_as_reported | 2.26 | R02 Table 1; secondary when other authors are named |
| Zhou2022_UB_as_reported | 2.34 | R02 Table 1; secondary when other authors are named |
| Zhou2022_SOCP | 2.3 | R02 Table 1; secondary when other authors are named |
| Wang2021_SOCP | 2.24 | R02 Table 1; secondary when other authors are named |
| Jia2024_eSPFEM | 2.27 | R02 Table 1; secondary when other authors are named |

補足仮定・制限:
- R02 uses explicit large-deformation SPFEM; compare failure onset, not its 6 s final displacement.
- No explicit tensile cap was specified in the inspected benchmark description; this input adds none.
- Fs<1 must be allowed; the independent-gravity trial profile is an explicit static-FEM adaptation.

### P05A — Three horizontal layers

- 出典箇所: p13 Fig10 Tables3–4
- 由来区分: `derived_associated_twin`。面積: **475 m²**。
- 外形（反時計回り）: `(0, 0) → (45, 0) → (45, 5) → (35, 5) → (15, 15) → (0, 15)`
- psi政策: `equal_reduced_phi`。各Fs独立自重試行を文献比較用の基本とする。
- 実解析/VBA: **NOT_RUN**。

| 材料 | E(kPa) | ν | γ(kN/m³) | c(kPa) | φ(°) | ψ初期(°) |
|---|---:|---:|---:|---:|---:|---:|
| upper | 60000 | 0.3 | 18 | 14 | 18 | 18 |
| middle | 80000 | 0.3 | 19.5 | 16.8 | 20 | 20 |
| base | 100000 | 0.3 | 21 | 1.9 | 23 | 23 |

領域:
- `upper`: `(0, 10) → (25, 10) → (15, 15) → (0, 15)`
- `middle`: `(0, 5) → (35, 5) → (25, 10) → (0, 10)`
- `base`: `(0, 0) → (45, 0) → (45, 5) → (0, 5)`

メッシュ目標h: M0=1.875m, M1=0.9375m, M2=0.46875m。

この派生入力/診断入力の定量Fsは未設定。元ケースから参照値をコピーしていない。

補足仮定・制限:
- R02 lists initial psi=9 degrees but its reduction policy is not fully specified here; N version explicitly holds psi=9.
- Slope ratio 1:2 is used exactly; the printed 26.57 degrees is rounded.

### P05N — Three horizontal layers

- 出典箇所: p13 Fig10 Tables3–4
- 由来区分: `literature_with_explicit_psi_policy`。面積: **475 m²**。
- 外形（反時計回り）: `(0, 0) → (45, 0) → (45, 5) → (35, 5) → (15, 15) → (0, 15)`
- psi政策: `hold_initial`。各Fs独立自重試行を文献比較用の基本とする。
- 実解析/VBA: **NOT_RUN**。

| 材料 | E(kPa) | ν | γ(kN/m³) | c(kPa) | φ(°) | ψ初期(°) |
|---|---:|---:|---:|---:|---:|---:|
| upper | 60000 | 0.3 | 18 | 14 | 18 | 9 |
| middle | 80000 | 0.3 | 19.5 | 16.8 | 20 | 9 |
| base | 100000 | 0.3 | 21 | 1.9 | 23 | 9 |

領域:
- `upper`: `(0, 10) → (25, 10) → (15, 15) → (0, 15)`
- `middle`: `(0, 5) → (35, 5) → (25, 10) → (0, 10)`
- `base`: `(0, 0) → (45, 0) → (45, 5) → (0, 5)`

メッシュ目標h: M0=1.875m, M1=0.9375m, M2=0.46875m。

| 報告方法 | Fs | 性格 |
|---|---:|---|
| LEM | 1.725 | R02 Table4 |
| FEM/SPH/XMPM | 1.7 | R02 Table4 |
| eSPFEM | 1.72 | R02 Table4 |

補足仮定・制限:
- R02 lists initial psi=9 degrees but its reduction policy is not fully specified here; N version explicitly holds psi=9.
- Slope ratio 1:2 is used exactly; the printed 26.57 degrees is rounded.

### P06A — Thin frictional weak band, 28 m domain

- 出典箇所: R02 p14 Fig12 Tables5–6; R10 printed p210 Fig1 mirrored x -> 28-x
- 由来区分: `cross_checked_geometry_with_unreported_tensile_policy`。面積: **302.5 m²**。
- 外形（反時計回り）: `(0, 0) → (28, 0) → (28, 5) → (23, 5) → (20, 8) → (8, 15) → (0, 15)`
- psi政策: `equal_reduced_phi`。各Fs独立自重試行を文献比較用の基本とする。
- 実解析/VBA: **NOT_RUN**。

| 材料 | E(kPa) | ν | γ(kN/m³) | c(kPa) | φ(°) | ψ初期(°) |
|---|---:|---:|---:|---:|---:|---:|
| upper | 14000 | 0.3 | 19 | 20 | 35 | 35 |
| band | 14000 | 0.3 | 19 | 0 | 25 | 25 |
| base | 14000 | 0.3 | 19 | 10 | 35 | 35 |

領域:
- `upper`: `(0, 10) → (20, 7.5) → (23, 5) → (20, 8) → (8, 15) → (0, 15)`
- `band`: `(0, 9.5) → (20, 7.1) → (23, 4.5) → (28, 5) → (23, 5) → (20, 7.5) → (0, 10)`
- `base`: `(0, 0) → (28, 0) → (28, 5) → (23, 4.5) → (20, 7.1) → (0, 9.5)`

メッシュ目標h: M0=1.875m, M1=0.9375m, M2=0.46875m。
薄層: at least 2/4/8 elements across interior band thickness for M0/M1/M2; grade towards daylight tip。先端では厚さが0へ向かうため局所gradingと境界適合を併用する。

この派生入力/診断入力の定量Fsは未設定。元ケースから参照値をコピーしていない。

補足仮定・制限:
- The lower band boundary continues from (23,4.5) to (28,5); the vendor coordinate figure resolves the faint end in R02.
- R10 vendor comparisons use model-specific tensile settings not fully stated by R02; none_explicit is an adopted policy, not an exact vendor reproduction.
- Do not force agreement with 0.90 by adding a cap or changing dilation silently.

### P06N — Thin frictional weak band, 28 m domain

- 出典箇所: R02 p14 Fig12 Tables5–6; R10 printed p210 Fig1 mirrored x -> 28-x
- 由来区分: `cross_checked_geometry_with_unreported_tensile_policy`。面積: **302.5 m²**。
- 外形（反時計回り）: `(0, 0) → (28, 0) → (28, 5) → (23, 5) → (20, 8) → (8, 15) → (0, 15)`
- psi政策: `hold_initial`。各Fs独立自重試行を文献比較用の基本とする。
- 実解析/VBA: **NOT_RUN**。

| 材料 | E(kPa) | ν | γ(kN/m³) | c(kPa) | φ(°) | ψ初期(°) |
|---|---:|---:|---:|---:|---:|---:|
| upper | 14000 | 0.3 | 19 | 20 | 35 | 0 |
| band | 14000 | 0.3 | 19 | 0 | 25 | 0 |
| base | 14000 | 0.3 | 19 | 10 | 35 | 0 |

領域:
- `upper`: `(0, 10) → (20, 7.5) → (23, 5) → (20, 8) → (8, 15) → (0, 15)`
- `band`: `(0, 9.5) → (20, 7.1) → (23, 4.5) → (28, 5) → (23, 5) → (20, 7.5) → (0, 10)`
- `base`: `(0, 0) → (28, 0) → (28, 5) → (23, 4.5) → (20, 7.1) → (0, 9.5)`

メッシュ目標h: M0=1.875m, M1=0.9375m, M2=0.46875m。
薄層: at least 2/4/8 elements across interior band thickness for M0/M1/M2; grade towards daylight tip。先端では厚さが0へ向かうため局所gradingと境界適合を併用する。

| 報告方法 | Fs | 性格 |
|---|---:|---|
| LEM | 0.927 | R02 Table6 |
| FEM | 0.86 | R02 Table6 |
| eSPFEM | 0.9 | R02 Table6 |

補足仮定・制限:
- The lower band boundary continues from (23,4.5) to (28,5); the vendor coordinate figure resolves the faint end in R02.
- R10 vendor comparisons use model-specific tensile settings not fully stated by R02; none_explicit is an adopted policy, not an exact vendor reproduction.
- Do not force agreement with 0.90 by adding a cap or changing dilation silently.

### P07A — Griffiths–Lane Example1, D=1

- 出典箇所: Example1, Fig1, Table2, Fig2
- 由来区分: `nondimensional_literature_case_dimensionalized`。面積: **220 m²**。
- 外形（反時計回り）: `(0, 0) → (32, 0) → (12, 10) → (0, 10)`
- psi政策: `equal_reduced_phi`。各Fs独立自重試行を文献比較用の基本とする。
- 実解析/VBA: **NOT_RUN**。

| 材料 | E(kPa) | ν | γ(kN/m³) | c(kPa) | φ(°) | ψ初期(°) |
|---|---:|---:|---:|---:|---:|---:|
| soil | 100000 | 0.3 | 20 | 10 | 20 | 20 |

領域:
- `soil`: `(0, 0) → (32, 0) → (12, 10) → (0, 10)`

メッシュ目標h: M0=1.25m, M1=0.625m, M2=0.3125m。

この派生入力/診断入力の定量Fsは未設定。元ケースから参照値をコピーしていない。

補足仮定・制限:
- H=10 m and gamma=20 kN/m3 are chosen dimensionalization; c/(gamma H)=0.05; crest width=1.2H; slope run=2H.
- E=100000 kPa and nu=0.3 are the nominal implementation values adopted here, not a claim of identical archived author input.
- R03 uses nonassociated psi=0; A is a separately labeled associated twin.
- No right roller is added: the toe lies at the fixed bottom endpoint.

### P07N — Griffiths–Lane Example1, D=1

- 出典箇所: Example1, Fig1, Table2, Fig2
- 由来区分: `nondimensional_literature_case_dimensionalized`。面積: **220 m²**。
- 外形（反時計回り）: `(0, 0) → (32, 0) → (12, 10) → (0, 10)`
- psi政策: `hold_initial`。各Fs独立自重試行を文献比較用の基本とする。
- 実解析/VBA: **NOT_RUN**。

| 材料 | E(kPa) | ν | γ(kN/m³) | c(kPa) | φ(°) | ψ初期(°) |
|---|---:|---:|---:|---:|---:|---:|
| soil | 100000 | 0.3 | 20 | 10 | 20 | 0 |

領域:
- `soil`: `(0, 0) → (32, 0) → (12, 10) → (0, 10)`

メッシュ目標h: M0=1.25m, M1=0.625m, M2=0.3125m。

| 報告方法 | Fs | 性格 |
|---|---:|---|
| FEM last reported converged | 1.35 | R03 Table2 |
| FEM first reported nonconverged at 1000 iterations | 1.4 | R03 Table2; not exact limit |
| Bishop-Morgenstern chart cross-check | 1.38 | R03 Fig2 |

補足仮定・制限:
- H=10 m and gamma=20 kN/m3 are chosen dimensionalization; c/(gamma H)=0.05; crest width=1.2H; slope run=2H.
- E=100000 kPa and nu=0.3 are the nominal implementation values adopted here, not a claim of identical archived author input.
- R03 uses nonassociated psi=0; A is a separately labeled associated twin.
- No right roller is added: the toe lies at the fixed bottom endpoint.

### P08A — Uniform undrained clay

- 出典箇所: Example4 Fig9; homogeneous limit also Example3
- 由来区分: `nondimensional_literature_case_dimensionalized`。面積: **900 m²**。
- 外形（反時計回り）: `(0, 0) → (60, 0) → (60, 10) → (40, 10) → (20, 20) → (0, 20)`
- psi政策: `equal_reduced_phi`。各Fs独立自重試行を文献比較用の基本とする。
- 実解析/VBA: **NOT_RUN**。

| 材料 | E(kPa) | ν | γ(kN/m³) | c(kPa) | φ(°) | ψ初期(°) |
|---|---:|---:|---:|---:|---:|---:|
| slope | 100000 | 0.3 | 20 | 50 | 0 | 0 |
| foundation | 100000 | 0.3 | 20 | 50 | 0 | 0 |

領域:
- `slope`: `(0, 10) → (40, 10) → (20, 20) → (0, 20)`
- `foundation`: `(0, 0) → (60, 0) → (60, 10) → (0, 10)`

メッシュ目標h: M0=2.5m, M1=1.25m, M2=0.625m。

| 報告方法 | Fs | 性格 |
|---|---:|---|
| Taylor homogeneous comparison | 1.47 | R03 Examples3/4; not exact finite-mesh value |

補足仮定・制限:
- H=10 m, gamma=20; cu1/(gamma H)=0.25; total-stress Tresca represented as MC phi=psi=0.
- No pore-pressure degree of freedom is added; E=100000 kPa, nu=0.3 are explicit numerical choices.
- For ratio 0.6 and 1.5 no digitized graph value is invented; establish the mesh-specific golden by independent solvers.

### P09A — Weak clay foundation

- 出典箇所: Example4 Fig9; homogeneous limit also Example3
- 由来区分: `nondimensional_literature_case_dimensionalized`。面積: **900 m²**。
- 外形（反時計回り）: `(0, 0) → (60, 0) → (60, 10) → (40, 10) → (20, 20) → (0, 20)`
- psi政策: `equal_reduced_phi`。各Fs独立自重試行を文献比較用の基本とする。
- 実解析/VBA: **NOT_RUN**。

| 材料 | E(kPa) | ν | γ(kN/m³) | c(kPa) | φ(°) | ψ初期(°) |
|---|---:|---:|---:|---:|---:|---:|
| slope | 100000 | 0.3 | 20 | 50 | 0 | 0 |
| foundation | 100000 | 0.3 | 20 | 30 | 0 | 0 |

領域:
- `slope`: `(0, 10) → (40, 10) → (20, 20) → (0, 20)`
- `foundation`: `(0, 0) → (60, 0) → (60, 10) → (0, 10)`

メッシュ目標h: M0=2.5m, M1=1.25m, M2=0.625m。

この派生入力/診断入力の定量Fsは未設定。元ケースから参照値をコピーしていない。

補足仮定・制限:
- H=10 m, gamma=20; cu1/(gamma H)=0.25; total-stress Tresca represented as MC phi=psi=0.
- No pore-pressure degree of freedom is added; E=100000 kPa, nu=0.3 are explicit numerical choices.
- For ratio 0.6 and 1.5 no digitized graph value is invented; establish the mesh-specific golden by independent solvers.

### P10A — Clay foundation mechanism transition

- 出典箇所: Example4 Fig9; homogeneous limit also Example3
- 由来区分: `nondimensional_literature_case_dimensionalized`。面積: **900 m²**。
- 外形（反時計回り）: `(0, 0) → (60, 0) → (60, 10) → (40, 10) → (20, 20) → (0, 20)`
- psi政策: `equal_reduced_phi`。各Fs独立自重試行を文献比較用の基本とする。
- 実解析/VBA: **NOT_RUN**。

| 材料 | E(kPa) | ν | γ(kN/m³) | c(kPa) | φ(°) | ψ初期(°) |
|---|---:|---:|---:|---:|---:|---:|
| slope | 100000 | 0.3 | 20 | 50 | 0 | 0 |
| foundation | 100000 | 0.3 | 20 | 75 | 0 | 0 |

領域:
- `slope`: `(0, 10) → (40, 10) → (20, 20) → (0, 20)`
- `foundation`: `(0, 0) → (60, 0) → (60, 10) → (0, 10)`

メッシュ目標h: M0=2.5m, M1=1.25m, M2=0.625m。

この派生入力/診断入力の定量Fsは未設定。元ケースから参照値をコピーしていない。

補足仮定・制限:
- H=10 m, gamma=20; cu1/(gamma H)=0.25; total-stress Tresca represented as MC phi=psi=0.
- No pore-pressure degree of freedom is added; E=100000 kPa, nu=0.3 are explicit numerical choices.
- For ratio 0.6 and 1.5 no digitized graph value is invented; establish the mesh-specific golden by independent solvers.

### D01A — All c=0, 45 degree face, phi=30

- 出典箇所: p8 Fig3 Table1
- 由来区分: `designed_diagnostic`。面積: **158 m²**。
- 外形（反時計回り）: `(0, 0) → (20, 0) → (20, 4) → (16, 4) → (10, 10) → (0, 10)`
- psi政策: `equal_reduced_phi`。各Fs独立自重試行を文献比較用の基本とする。
- 実解析/VBA: **NOT_RUN**。

| 材料 | E(kPa) | ν | γ(kN/m³) | c(kPa) | φ(°) | ψ初期(°) |
|---|---:|---:|---:|---:|---:|---:|
| soil | 14000 | 0.3 | 20 | 0 | 30 | 30 |

領域:
- `soil`: `(0, 0) → (20, 0) → (20, 4) → (16, 4) → (10, 10) → (0, 10)`

メッシュ目標h: M0=1.25m, M1=0.625m, M2=0.3125m。

この派生入力/診断入力の定量Fsは未設定。元ケースから参照値をコピーしていない。

補足仮定・制限:
- R02 uses explicit large-deformation SPFEM; compare failure onset, not its 6 s final displacement.
- No explicit tensile cap was specified in the inspected benchmark description; this input adds none.
- Fs<1 must be allowed; the independent-gravity trial profile is an explicit static-FEM adaptation.

### D02A — All c=0, slope ratio 2:3, phi=45

- 出典箇所: 設計診断。文献再現ではない。
- 由来区分: `designed_diagnostic`。面積: **1400 m²**。
- 外形（反時計回り）: `(0, 0) → (70, 0) → (70, 10) → (50, 10) → (20, 30) → (0, 30)`
- psi政策: `equal_reduced_phi`。各Fs独立自重試行を文献比較用の基本とする。
- 実解析/VBA: **NOT_RUN**。

| 材料 | E(kPa) | ν | γ(kN/m³) | c(kPa) | φ(°) | ψ初期(°) |
|---|---:|---:|---:|---:|---:|---:|
| soil | 30000 | 0.3 | 20 | 0 | 45 | 45 |

領域:
- `soil`: `(0, 0) → (70, 0) → (70, 10) → (50, 10) → (20, 30) → (0, 30)`

メッシュ目標h: M0=3.75m, M1=1.875m, M2=0.9375m。

この派生入力/診断入力の定量Fsは未設定。元ケースから参照値をコピーしていない。

補足仮定・制限:
- Designed case, not reconstruction of the saved G1 workbook.

tan45°/tan(arctan(2/3))=1.5は表層滑りの極限機構の照合値。有限メッシュのFs=1.5を保証するものではない。保存済みG1モデルを復元した入力でもない。

### D03N — D02 geometry, c=0 nonassociated psi=0

- 出典箇所: 設計診断。文献再現ではない。
- 由来区分: `designed_diagnostic`。面積: **1400 m²**。
- 外形（反時計回り）: `(0, 0) → (70, 0) → (70, 10) → (50, 10) → (20, 30) → (0, 30)`
- psi政策: `hold_initial`。各Fs独立自重試行を文献比較用の基本とする。
- 実解析/VBA: **NOT_RUN**。

| 材料 | E(kPa) | ν | γ(kN/m³) | c(kPa) | φ(°) | ψ初期(°) |
|---|---:|---:|---:|---:|---:|---:|
| soil | 30000 | 0.3 | 20 | 0 | 45 | 0 |

領域:
- `soil`: `(0, 0) → (70, 0) → (70, 10) → (50, 10) → (20, 30) → (0, 30)`

メッシュ目標h: M0=3.75m, M1=1.875m, M2=0.9375m。

この派生入力/診断入力の定量Fsは未設定。元ケースから参照値をコピーしていない。

補足仮定・制限:
- Designed case, not reconstruction of the saved G1 workbook.

### D04A — Elastic modulus sensitivity E/10

- 出典箇所: p8 Fig3 Table1
- 由来区分: `designed_diagnostic`。面積: **158 m²**。
- 外形（反時計回り）: `(0, 0) → (20, 0) → (20, 4) → (16, 4) → (10, 10) → (0, 10)`
- psi政策: `equal_reduced_phi`。各Fs独立自重試行を文献比較用の基本とする。
- 実解析/VBA: **NOT_RUN**。

| 材料 | E(kPa) | ν | γ(kN/m³) | c(kPa) | φ(°) | ψ初期(°) |
|---|---:|---:|---:|---:|---:|---:|
| soil | 1400 | 0.3 | 20 | 10 | 25 | 25 |

領域:
- `soil`: `(0, 0) → (20, 0) → (20, 4) → (16, 4) → (10, 10) → (0, 10)`

メッシュ目標h: M0=1.25m, M1=0.625m, M2=0.3125m。

この派生入力/診断入力の定量Fsは未設定。元ケースから参照値をコピーしていない。

補足仮定・制限:
- R02 uses explicit large-deformation SPFEM; compare failure onset, not its 6 s final displacement.
- No explicit tensile cap was specified in the inspected benchmark description; this input adds none.
- Fs<1 must be allowed; the independent-gravity trial profile is an explicit static-FEM adaptation.

### D05A — Geometric similitude: L x1000, gamma /1000

- 出典箇所: p8 Fig3 Table1
- 由来区分: `designed_diagnostic`。面積: **1.58e+08 m²**。
- 外形（反時計回り）: `(0, 0) → (20000, 0) → (20000, 4000) → (16000, 4000) → (10000, 10000) → (0, 10000)`
- psi政策: `equal_reduced_phi`。各Fs独立自重試行を文献比較用の基本とする。
- 実解析/VBA: **NOT_RUN**。

| 材料 | E(kPa) | ν | γ(kN/m³) | c(kPa) | φ(°) | ψ初期(°) |
|---|---:|---:|---:|---:|---:|---:|
| soil | 14000 | 0.3 | 0.02 | 10 | 25 | 25 |

領域:
- `soil`: `(0, 0) → (20000, 0) → (20000, 4000) → (16000, 4000) → (10000, 10000) → (0, 10000)`

メッシュ目標h: M0=1250m, M1=625m, M2=312.5m。

この派生入力/診断入力の定量Fsは未設定。元ケースから参照値をコピーしていない。

補足仮定・制限:
- R02 uses explicit large-deformation SPFEM; compare failure onset, not its 6 s final displacement.
- No explicit tensile cap was specified in the inspected benchmark description; this input adds none.
- Fs<1 must be allowed; the independent-gravity trial profile is an explicit static-FEM adaptation.

不変性試験: same ideal Fs as P02A; displacements divided by length should match under same discretization

### D06A — Translation covariance (+1000,+2000)

- 出典箇所: p8 Fig3 Table1
- 由来区分: `designed_diagnostic`。面積: **158 m²**。
- 外形（反時計回り）: `(1000, 2000) → (1020, 2000) → (1020, 2004) → (1016, 2004) → (1010, 2010) → (1000, 2010)`
- psi政策: `equal_reduced_phi`。各Fs独立自重試行を文献比較用の基本とする。
- 実解析/VBA: **NOT_RUN**。

| 材料 | E(kPa) | ν | γ(kN/m³) | c(kPa) | φ(°) | ψ初期(°) |
|---|---:|---:|---:|---:|---:|---:|
| soil | 14000 | 0.3 | 20 | 10 | 25 | 25 |

領域:
- `soil`: `(1000, 2000) → (1020, 2000) → (1020, 2004) → (1016, 2004) → (1010, 2010) → (1000, 2010)`

メッシュ目標h: M0=1.25m, M1=0.625m, M2=0.3125m。

この派生入力/診断入力の定量Fsは未設定。元ケースから参照値をコピーしていない。

補足仮定・制限:
- R02 uses explicit large-deformation SPFEM; compare failure onset, not its 6 s final displacement.
- No explicit tensile cap was specified in the inspected benchmark description; this input adds none.
- Fs<1 must be allowed; the independent-gravity trial profile is an explicit static-FEM adaptation.

不変性試験: same Fs and displacement field translated back as P02A

### D07A — Crest surcharge 20 kPa on x=2..8

- 出典箇所: p8 Fig3 Table1
- 由来区分: `designed_diagnostic`。面積: **158 m²**。
- 外形（反時計回り）: `(0, 0) → (20, 0) → (20, 4) → (16, 4) → (10, 10) → (0, 10)`
- psi政策: `equal_reduced_phi`。各Fs独立自重試行を文献比較用の基本とする。
- 実解析/VBA: **NOT_RUN**。

| 材料 | E(kPa) | ν | γ(kN/m³) | c(kPa) | φ(°) | ψ初期(°) |
|---|---:|---:|---:|---:|---:|---:|
| soil | 14000 | 0.3 | 20 | 10 | 25 | 25 |

領域:
- `soil`: `(0, 0) → (20, 0) → (20, 4) → (16, 4) → (10, 10) → (0, 10)`

メッシュ目標h: M0=1.25m, M1=0.625m, M2=0.3125m。

この派生入力/診断入力の定量Fsは未設定。元ケースから参照値をコピーしていない。

追加荷重: `[{"type": "uniform_boundary_traction", "a": [2, 10], "b": [8, 10], "traction_kPa": [0, -20]}]`

補足仮定・制限:
- R02 uses explicit large-deformation SPFEM; compare failure onset, not its 6 s final displacement.
- No explicit tensile cap was specified in the inspected benchmark description; this input adds none.
- Fs<1 must be allowed; the independent-gravity trial profile is an explicit static-FEM adaptation.

### D08A — Fixed pore pressure field, water elevation y=3

- 出典箇所: p8 Fig3 Table1
- 由来区分: `designed_diagnostic`。面積: **158 m²**。
- 外形（反時計回り）: `(0, 0) → (20, 0) → (20, 4) → (16, 4) → (10, 10) → (0, 10)`
- psi政策: `equal_reduced_phi`。各Fs独立自重試行を文献比較用の基本とする。
- 実解析/VBA: **NOT_RUN**。

| 材料 | E(kPa) | ν | γ(kN/m³) | c(kPa) | φ(°) | ψ初期(°) |
|---|---:|---:|---:|---:|---:|---:|
| soil | 14000 | 0.3 | 20 | 10 | 25 | 25 |

領域:
- `soil`: `(0, 0) → (20, 0) → (20, 4) → (16, 4) → (10, 10) → (0, 10)`

メッシュ目標h: M0=1.25m, M1=0.625m, M2=0.3125m。

この派生入力/診断入力の定量Fsは未設定。元ケースから参照値をコピーしていない。

間隙水圧: `{"mode": "fixed_hydrostatic", "water_elevation_m": 3.0, "gamma_w_kPa_m": 9.81, "formula": "9.81*max(3-y,0)", "surface_water_traction": "none; exposed ground is above y=3"}`

補足仮定・制限:
- R02 uses explicit large-deformation SPFEM; compare failure onset, not its 6 s final displacement.
- No explicit tensile cap was specified in the inspected benchmark description; this input adds none.
- Fs<1 must be allowed; the independent-gravity trial profile is an explicit static-FEM adaptation.

### D09A — Near-incompressibility diagnostic nu=0.49

- 出典箇所: p8 Fig3 Table1
- 由来区分: `designed_diagnostic`。面積: **158 m²**。
- 外形（反時計回り）: `(0, 0) → (20, 0) → (20, 4) → (16, 4) → (10, 10) → (0, 10)`
- psi政策: `equal_reduced_phi`。各Fs独立自重試行を文献比較用の基本とする。
- 実解析/VBA: **NOT_RUN**。

| 材料 | E(kPa) | ν | γ(kN/m³) | c(kPa) | φ(°) | ψ初期(°) |
|---|---:|---:|---:|---:|---:|---:|
| soil | 14000 | 0.49 | 20 | 10 | 25 | 25 |

領域:
- `soil`: `(0, 0) → (20, 0) → (20, 4) → (16, 4) → (10, 10) → (0, 10)`

メッシュ目標h: M0=1.25m, M1=0.625m, M2=0.3125m。

この派生入力/診断入力の定量Fsは未設定。元ケースから参照値をコピーしていない。

補足仮定・制限:
- R02 uses explicit large-deformation SPFEM; compare failure onset, not its 6 s final displacement.
- No explicit tensile cap was specified in the inspected benchmark description; this input adds none.
- Fs<1 must be allowed; the independent-gravity trial profile is an explicit static-FEM adaptation.

## メッシュと基準解を作る具体的手順

1. 中立JSON/CSVから既存入力へ変換し、各領域の面積・材料・外周と総重量を照合する。
2. 同一メッシャーでM0/M1/M2を生成し、節点・要素・材料ID・拘束・荷重・積分点を保存する。Q4/Q8を途中で交換しない。
3. 領域界面に適合した要素を使う。P06の薄層内の横断要素数を数え、重心割当だけで薄層を欠落させない。
4. まず基準ソルバーでF_s区間と固定Fsの出力を採る。FEM収束率が不足しているケースでは高速化比較を保留する。
5. M1/M2の差、必要なら境界拡張も確認する。領域拡張したものは別ケースであって、原文献座標を無断更新しない。
6. 確定した基準ログをgoldenへ保存して初めて新方式と比較する。

## ファイル形式の注意

CSVの外形/領域/境界/材料はExcelで読みやすい補助表。D07の追加荷重、D08の水圧、初期化政策等はJSONが原本であり、材料CSVだけでは全条件を表現しない。要素接続データは未生成。既存コードへ接続するadapterと実メッシュ生成はE0/E1で行う。
