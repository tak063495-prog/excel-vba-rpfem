# D01A ログ要約（G1R19 2%プリセット再実行・Rico）

- 開始 16:00:52 JST / RPX_Stop 送信 17:30:56 JST（制限 5,400 s = 前回 native26 の D01A 1,787.2 s ×3）
- 結果: TIMEOUT / ANALYSIS_CANCELLED（RPX_Stop による協調停止。Excel 強制終了なし）、Fs 未確定
- 停止時点: 下界 ≈0.577 / 上界 ≈0.700、GAP ≈19.2%。精度回復5回で GAP はほぼ不変（19.27%→19.22%）
- 経過秒 = analysis.tsv の elapsed_s（解析開始からの秒）。analysis.tsv 原本は UTF-16LE、analysis_utf8.tsv は UTF-8 変換版

## 段階の開始（analysis.tsv 初出）

| 段階 | 開始 s |
|---|---:|
| coarse_0 | 1 |
| adapt_1_band | 247 |
| band_refine | 454 |
| final | 706 |
| accuracy_1 | 918 |
| accuracy_2 | 1169 |
| accuracy_3 | 1701 |
| accuracy_4 | 1873 |
| accuracy_5 | 2899 |
| lower_angular_fan_0 | 3452 |

## 主要イベント（analysis.tsv 原文抜粋）

| s | 段階 | イベント | 内容 |
|---:|---|---|---|
| 1.1 |  | zero_stress_fan_repair | `split_count=7;old_elements=355;new_elements=369;scope=GEOMETRY_ONLY;loads=raw;constitutive_unchanged=True` |
| 246.5 | coarse_0 | adapt_cycle | `cycle=0;route=conventional;elements=369;nodes=214;lower=0.577148319824458;upper=0.700195479003999;gap=0.192660988047849;fs_calls=24` |
| 450.8 | adapt_1_band | adapt_cycle | `cycle=1;route=conventional;elements=369;nodes=214;lower=0.577148319824458;upper=0.700195479003999;gap=0.192660988047849;gain=0;fs_calls=19` |
| 451.4 | adapt_1_band | zero_stress_fan_repair | `split_count=1;old_elements=1156;new_elements=1158;scope=GEOMETRY_ONLY;loads=raw;constitutive_unchanged=True` |
| 452.2 | adapt_1_band | zero_stress_fan_repair | `split_count=1;old_elements=869;new_elements=871;scope=GEOMETRY_ONLY;loads=raw;constitutive_unchanged=True` |
| 916.5 | final | zero_stress_fan_repair | `split_count=1;old_elements=1156;new_elements=1158;scope=GEOMETRY_ONLY;loads=raw;constitutive_unchanged=True` |
| 917.3 | final | zero_stress_fan_repair | `split_count=1;old_elements=869;new_elements=871;scope=GEOMETRY_ONLY;loads=raw;constitutive_unchanged=True` |
| 1165.7 | accuracy_1 | accuracy_round | `round=1;elements=531;gap=0.19226076059986;gain=6.81244907847533E-05;stalls=0` |
| 1166.5 | accuracy_1 | zero_stress_fan_repair | `split_count=1;old_elements=1614;new_elements=1616;scope=GEOMETRY_ONLY;loads=raw;constitutive_unchanged=True` |
| 1698.3 | accuracy_2 | accuracy_round | `round=2;elements=531;gap=0.19226076059986;gain=0;stalls=1` |
| 1699.1 | accuracy_2 | zero_stress_fan_repair | `split_count=1;old_elements=1614;new_elements=1616;scope=GEOMETRY_ONLY;loads=raw;constitutive_unchanged=True` |
| 1871.5 | accuracy_3 | accuracy_round | `round=3;elements=697;gap=0.192248179949237;gain=6.54353524025789E-05;stalls=0` |
| 2897.0 | accuracy_4 | accuracy_round | `round=4;elements=697;gap=0.192248179949237;gain=0;stalls=1` |
| 3314.6 | accuracy_5 | fs_numerical_unknown | `reason=HSD_DD_PRODUCT_UNDERFLOW;endpoints_unchanged=True;phase=fs_trial` |
| 3320.2 | accuracy_5 | fs_numerical_unknown | `reason=HSD_DD_PRODUCT_UNDERFLOW;endpoints_unchanged=True;phase=fs_trial` |
| 3449.0 | accuracy_5 | accuracy_round | `round=5;elements=697;gap=0.192248179949237;gain=0;stalls=2` |
| 3449.0 | accuracy_5 | accuracy_stop | `reason=stalled_after_indicator_retry;round=5` |
| 3449.0 | accuracy_5 | accuracy_end | `attempted_rounds=5;elements=697;gap=0.192248179949237` |
| 3451.4 | accuracy_5 | zero_stress_fan_repair | `split_count=7;old_elements=4002;new_elements=4016;scope=GEOMETRY_ONLY;loads=raw;constitutive_unchanged=True` |
| 3452.2 | accuracy_5 | lower_uniform_candidate | `target=4000;elements=4016;upper_mesh_unchanged=1;certificate=0` |
| 5399.7 | lower_angular_fan_0 | solve_end | `status=CANCELLED;wall_seconds=100.691247899988;last_iteration_index=3;objective=0;primal=0.31409193544465;dual=0.314091935401261;gap=0;message=ANALYSIS_CANCELLED` |
| 5399.8 | lower_angular_fan_0 | error_origin | `number=18;source=VBAProject;line=0;stage=control_analysis;message=ANALYSIS_CANCELLED` |
| 5399.8 | lower_angular_fan_0 | run_end | `status=CANCELLED;solve_count=208;sampled_peak_private_bytes=1344311296;logging_seconds_before_run_end=26.91005340009;message=ANALYSIS_CANCELLED` |

## stage.txt 末尾（停止直前・原文。末尾数値は Excel Timer＝0時からの秒、63051≈17:30:51）

```
# HSD 17 residual=9.1E-03 gap=3.6E+05 62713.28
# HSD 18 residual=9.1E-03 gap=2.2E+06 62738.01
# HSD 19 residual=9.1E-03 gap=1.0E+08 62833.44
# lower F=0.574271076584141 62956.62
# HSD 0 residual=1.0E+00 gap=2.4E+04 63031.7
# HSD 1 residual=6.2E-01 gap=1.5E+04 63038.34
# HSD 2 residual=4.7E-01 gap=1.1E+04 63044.98
# HSD 3 residual=2.1E-01 gap=5.1E+03 63051.52
```
