# 実 VBA 検証ハンドオフ

新 source1 / import CP932 / 全 G1R20 依存ソース / DLL と source manifest を凍結した所有ブックで行います。旧 TerminalRoot・UpperRecovery・Combined・証拠は変更しません。

1. compile/reopen、既存 UpperRecovery 純配列10検査、前 sibling の terminal gate 無変更契約を再実行します。
2. test-only bridge で小さな原 program / dual を用意します。原 CurrentProgram を満たす fresh compile を用い、q8/Fs/model/policy/epoch/cost を固定します。等式の自由 y だけを意図的に乱した candidate を入力し、補正前が raw stat 門不合格、補正後が全元門合格、cone z の uint64 bit 不変、source array 不変、1回200上限・optimizer0を確認します。cone-invalid / separation-invalid / lower / q4 / wrong layout / nonfinite は不採用です。
3. Python 最終 N source と同じ元式を literal CSR と全 dual に export し、`tests/native_lsmr_probe.py` 相当の同 DLL API を別 Python process で呼びます。native CSR と Python モデル再組立を区別して報告し、各原 residual と補正後全ベクトルを比較します。
4. `dual_repair_before_lsmr` に既存 TestHook の取消18・入力変更を注入し、APIへ到達せず解析中断することを確認します。API後の Fs/q/epoch/model/policy/P4Solve/candidate ID 変更も拒否、Direct Restore 後に入力変更を HSD fallbackへ変換しないことを確認します。
5. 新 helper regularization は flag0 ではP0、flag1 upperだけP1e-8、原元コスト/constraints bytes同一、下界P0を確認します。全 dual stage 証拠を diagnostic と accepted に分け、z変更・二度補正・QP dualの誤流入を拒否します。

この手順は実施結果ではありません。10秒は API後の採否上限で、同期APIの中断保証ではありません。外部所有PID watchdogを継続します。
