# 原 dual の自由な等式成分だけを1回補正する隔離試作

UpperRecovery v1 の凍結 Native を基準とする1モジュールです。既存 DLL を変更せず `rpfem_accuracy_equalities` を利用します。同 API の Rust `equality.rs` は SciPy 1.17 の real undamped LSMR 移植です。原 LP / 原 dual feasibility の候補だけを扱い、cap-QP の dual は扱いません。

元の符号規約は `cost + E.T*y - G.T*z = 0`、VBA の CSR は `[E;-G]` です。元 CSR の等式部分だけを転置した CSR を作り、`E.T*dy = -(cost + E.T*y - G.T*z)` を200反復 / atol=btol=1e-15 / conlim=1e12 で1回解きます。`y+=dy` のみを行い、cone 変数 z は有限値の等値と +/-0 の bit 比較で厳密に不変確認します。元 DualAudit の閾値と separation を全て再確認し、LSMR の status は採否に使いません。API 実行が10秒を超えた候補も採用しません。API は同期処理なので、10秒は実行後の採否上限であり強制中断保証ではありません。総 watchdog は別です。

`ACCURACY_DLL_DUAL_LINEAR_REPAIR=0` が既定で、全下界も従来経路です。1 にした上界の original LP dual または original-dual-feasibility primal が原門を通らなかった場合だけ補正を試します。source candidate は不採用・例外時に変更しません。入力/context/取消は前後で照合し、Direct の Restore 後も例外を伝播します。補正から primal 場や physical Fs bound は作らず、既存の original dual guide 採用と根区間契約を通ります。

N の最終 Python 成功条件に合わせ、同 flag1 の上界 helper だけ `ACCURACY_DLL_DUAL_REGULARIZATION` を補助 Pdiag に設定できます。既定0、明示 G1R20 profile1e-8 / DUAL_SECONDS60 / DUAL_EQUILIBRATE1です。原物理目的と全原式、dual feasibility の制約構造は変更しません。primal は既存 EQUILIBRATE0 のままです。

`ACCURACY_DLL_CAPTURE=1` の場合、補正前後の全原プログラムと全 dual vector を `.dual_repair_before_*.bin` / `.dual_repair_after_*.bin` と context へ保存します。内容は AC18P001 の **dual layout** ですが、不採用 source も含むため通常の accepted `.accuracy_*.bin` 群から区別しています。ログに before/after 原監査、採否、LSMR反復・秒数・z不変を残します。これは diagnostic candidate で、保存しただけで原門に合格したことにはなりません。採用後には既存の通常 accepted proof も別保存されます。

Python 正本は `validation/G1R19_python26_dual_linear_repair_20261016/repair_original_dual.py`。N の source stat1.7849141851e-7 を一度の5反復で8.9245876586e-8へ低減し、cone5.6613868866e-13とzを維持、objective1.0000030000010307を原門で確認した経路です。VBA compile と実ブック上の同一性は別 micro で確認します。本フォルダの静的/FFI検査をその代わりとはしません。

solver0 の同 DLL FFI 検証は `results/ffi_20261016_v1/RESULTS.json` に保存しています。E は1210等式 /104868変数/28506非零で、既存66da DLL は5反復/.0076654秒、SciPy補正との差最大1.19e-22でした。元全dual門を通り、z の全 uint64 bit と source/DLL/mesh bytes を保存しました。

さらに `tests/vba_rhs_probe.py` は VBA と同じ compensated `[E;-G]` 残差を作って同 DLL へ渡す独立比較です。`results/ffi_compensated_20261016_v1/RESULTS.json` は5反復/.0077712秒、補正後 compensated stat8.9245876655e-8 / cone5.6613868866e-13 / objective1.0000030000010307で全門 PASS です。numpy 残差との差最大4.70e-15も記録しています。ここでも最適化求解は0、実 VBA 実行は0です。
