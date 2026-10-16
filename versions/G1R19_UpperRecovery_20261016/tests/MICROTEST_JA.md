# integration への実 Excel 検証手順（製品に試験コードを含めない）

所有 Excel / full source / 入力 / DLL の SHA と起動 PID・出生時刻・全コマンドを別出力へ凍結し、総時間 watchdog を付けます。既存 Combined10 / TerminalRoot / 凍結ブックを上書きしません。

1. 新 Native と既存 G1R20 依存モジュールで compile/reopen します。`tests/UpperRecoveryArrayProbe.bas` の Attribute/Option を除いた bridge を **試験コピーの Native 内**へ追記し、`RPX_TestUpperWorkArrays` を実行します。有限正仕事の式と9種類の拒否、元候補不変を確認します。solver0、証明0です。
2. fixed-Fs P06 上界を fresh assemble し、既存公開 `RPX_AccuracyRecoverPrimal` を呼びます。test-only source の RPX_AccuracySolve 返却直後と AuditUpperRecenter 後に LP/QP candidate・workRatio・元 CSR を export します。各 solve 前後の context と元 CSR bytes 同一、最大 LP1+QP2、同じ Fs、2回目 center が1回目の正規化済み rejected candidate であることを確認します。同一 literal LP/QP を Python FFI に渡す比較と、Python 元の affine LP からの違いは区別します。
3. 原 LP で既に元門を通るケースは QP0、QP1成功なら QP1、QP1不採用/2成功なら QP2、両方不採用なら False / recovered未割当を test-only solve stub でも確認します。SOLVED 等 status のみ成功し原 raw 門を破る候補は必ず False、modified dual を original guide にしないことを確認します。
4. `upper_recenter_lp_before` / `upper_recenter_qp_before` に既存 TestHook で入力変更・取消18を注入します。solve 前の再 Assert が入力変更を拒否、取消を呼出元へ伝播、live xx/objective/slack/field gate/diagnostic audit/root/証拠を保護することを確認します。P4Solve/candidate ID/Fs/epoch/q/policy/model 変更は各 solve/audit 前後で必ず停止することも試験コピーで確認します。
5. flag0 上界 / flag1 下界を同じ fresh 入力で比較し、旧処理本文と実経路が維持されることを確認します。ordinary default の設定行、閾値、物理式、根幅を変えません。

この手順は実施結果ではありません。ソース契約テストの PASS を実 VBA compile や実 P06 証明と混同しません。
