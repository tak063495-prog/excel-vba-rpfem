# 新規求解24ケースの部分証拠（2026-10-16）

P06以外の24ケースについて、今回新しく作った上下限の保存場を同梱する。
過去16ペアの再監査ではない。P06はこの束では未達であり、全26合格・正式リリースとはしない。

元q8物理場・両自己根・対称GAP2%の確認は24/26、正常cold実行の記録は20/26。
D01A/D05A/D08Aの別版実験はworker終了コード未記録、P04Nはdriver時間切れ。
その区別を各caseとmanifestに残した。

EVIDENCE.zip: 10,255,703 bytes、SHA256 318ae43a5502388f2490eaa1f11fa5341f50a5bcd12ee563be9a15dfe333ba89。
新しい別フォルダに展開し、Python環境を用意して `python REPLAY.py --output NEW_REPLAY_DIRECTORY` を実行する。
検証環境はPython3.13.14、NumPy2.3.5、SciPy1.17.0、Clarabel0.11.1。
元JSONの絶対パスは計算記録として変更せず、再生はmanifestの同梱相対フォルダを使用する。

独立した展開先・新規プロセスで24/24を再生し、元式・両完了根・2%を再確認した。
再生87.0887秒、solver呼出0、returncode0、全478memberと元ZIPのSHA不変。
再生時間は初期状態からの求解時間ではない。詳細はREPLAY_VERIFICATION.json。
