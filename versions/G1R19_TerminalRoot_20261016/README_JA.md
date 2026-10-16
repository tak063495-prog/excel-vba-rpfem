# 狭い根区間の試行順序（隔離sibling）

凍結Combined10を基準に、FsとAccuracyNativeの2モジュールだけを別フォルダで変更した。既存ブック・ソース・場・根・DLLは変更しない。UIはG1R20 integration担当が `SETTINGS_CONTRACT.json` に従って追加する。

通常は `FS_TERMINAL_SCHEDULE=0`。明示的な改善profileの1、q8、Accuracy有効、現在の根幅が既存tauを超え2.5tau以下のときだけ新順序を使う。tauは `FS_TOL*max(1,(a+b)/2)`。これは試行位置の選択条件であり、根を合格とする幅は従来のtauのまま。

下界は元区間の30%点、70%点を順に既存TryFsTrialPointで評価し、validな符号だけで端点を更新する。1点目の更新後に区間外となる2点目は評価しない。両点がUNKNOWNなら端点を維持し未完了で終了する。外側40反復は増やさない。候補の時間やsolver statusを根の符号に置き換えない。

上界ではTryFsValueがexact Fs/model/input epoch/q/policy/load phase/solve epoch/candidate IDに結び付けたone-shot requestをarmする。DirectBoundは入口で一度だけconsume/resetし、原問題exportと作業状態保存の後、既存の原dual feasibility helperを先に試す。成功は既存original_dual_helperのRestore/commit経路を共用し、guideを物理場へ採用しない。helperがFalseなら通常の原LPへ進み、同じDirectBound内ではhelperを二度試さない。取消は伝播し、return/errorでもrequestをresetする。FsBracketResetでも失効する。

G1R20改善profileのLP30秒・dual30秒は**別々の宣言予算**。helper先行が失敗した点では合計最大60秒と従来fallbackがあり得る。全体実行時間の外部監視は別。通常既定の予算を黙って変更せず、時間内終了と物理証拠を別に記録する。

原物理式・objective・q・Primal/DualAudit・FieldGate・ExactModel/epoch照合・分離条件・TryFsTrialPoint・SaveBracket・ActualRootBracket・GAP/根幅許容条件は維持する。実Excel compile、micro、P06数値結果は未確認。このsibling単独でP06達成やリリースを主張しない。
