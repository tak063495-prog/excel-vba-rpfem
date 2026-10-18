# G1R20 Fix7 rc1 検証（Winmax2、2026-10-18）

rc1（上界 conic を1回の失敗で省略）の結果です。SOURCE_COMMIT は `cac3203` です。
P10A 1793.5 秒、D05A 260.7 秒、P02A 316.1 秒、P05A 1387.6 秒で、Fs はすべて Fix6 と同じでした。
P05A は多くの点で conic が成功するので、rc2 では「2回以上失敗し、成功0回」のときだけ省略するように改めました。
作業用ブック（rc1 の xlsm）は容量のため含めていません。`delivery_rc1/WORKBOOK_VALIDATION.json` が残っています。
