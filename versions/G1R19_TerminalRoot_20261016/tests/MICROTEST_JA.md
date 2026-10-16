# Integrationへのmicro引継ぎ

Combined10を直接編集せず、新しい使い捨てG1R20候補へこの2moduleを上書きする。source/import SHAと所有PID/birth/full commandを凍結し、compile/save/reopen/compileを確認する。UIと全体900秒等のwatchdogはintegration/rootの契約を使う。

1. 通常Setup/default0では従来ルート。明示改善profileだけflag1、LP30/dual30、q8・from-start・GAP.02・FS_TOL1e-4。材料・水理・境界・policyの値と型が不変であることを比較する。
2. test-copyにだけprivate access wrapperを追加し、UpperDualFirstArm→Consumeを同一contextで一度True/二度目Falseと確認する。異なるFs、mode、model、epoch、q、policy、phase、solve epoch、candidate ID、reset後ではFalse。one-shot flagをguideReady/PhysicalGateAccepted等の証拠フラグとして使わない。
3. helper-firstをtest-only stubでFalseにし、原LP入口へ一度だけ進むことを確認する。helperの直後に取消18を注入し、原LP/HSDに進まず伝播することを確認する。input mutation/context mismatchも採用せず中断する。stubは製品ソースへ入れず、全指標が保存された物理場を装う値を生成しない。
4. lower terminalの試行をtest-copy内の既存TestTrialフックで制御し、UNKNOWN/UNKNOWNで端点・保存場・statusを不変かつ未完了、positive/negativeで正しい狭い区間、1点目negative後の70%点が更新区間外ならskip、30%点だけで幅達成時は70%点を求解しないことを確認する。各採否は既存TryFsTrialPointの全門へ戻す。
5. 新規tiny physical upper（既存D02等の同mesh、既知正符号側Fs）のhelper先行を一度実行し、原AC18dual全stationarity/cone/separation/context門とguide-only commitを独立再生する。反対側や無効helperは原LPへのfallbackを確認する。solver statusのみでPASSとしない。
6. 最後に製品2moduleだけの候補でP06をcold/明示continuation区別付きで評価し、途中のmesh/候補/根点列もPythonの成功手順と比較する。旧場を新meshへそのまま保存しない。失敗・未完了を含む全試行を残す。

静的契約テストは原門とcommit branchが基準とbyte正規化一致することを確認するが、実Excelの制御フロー試験の代わりにはならない。
