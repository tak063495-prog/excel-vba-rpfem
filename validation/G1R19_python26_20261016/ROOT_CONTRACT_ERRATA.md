# Fs根幅のメタデータ表記訂正（2026-10-16）

既存G1R19と再利用PythonコードのFs根完了条件は、絶対幅だけの1e-4ではなく、次の正規化条件である。

```text
0 < lo < hi
(hi - lo) / max(1, (hi + lo)/2) <= 1e-4
```

両端の原式に基づく符号 `[1, -1]`、case/mode/q/policy/physical-model/own-meshの同一性、下界ではlo・上界ではhiと保存Fsの一致も要求する。`python/accuracy_driver.py` の search と `python/report_accuracy.py` の root_is_complete はこの同じ条件を実装している。Fs>1では許される絶対幅が1e-4より大きくなるが、既存の根契約どおりであり、求解・物理監査条件の緩和ではない。

凍結した INPUTS_26.json の `acceptance.fs_root_tolerance_absolute` と `execution_profiles.fresh_Python26_reference.fs_root_tolerance_absolute` はメタデータの名称が不正確だった。値1e-4は上記正規化条件の係数を意味する。README、HISTORY_AUDIT、audit_case出力等の「根幅1e-4」もこの正規化条件として読む。ユーザーの新しい絶対1e-4要求が加わったという意味ではない。

実行中のソース・入力・ハッシュは変更しない。終了後に今回の入力スナップショットとSHA-256を保全した上で、正本のキーを `fs_root_relative_tolerance` に直し、正規化式を併記する。今回の入力SHAは `971bf1c476319e11361f35590c5a3ed6017c46adbb0f211b8feaa7b915b78e0f` である。

上下界の対称幅 `2*(U-L)/(U+L) <= .02`、q8、元の物理場と原双対の監査、原物理残差1e-7、原双対stationarity1e-7/cone1e-8は変更しない。小数点以下2桁の四捨五入一致は別の表示指標であり、2%判定や根完了条件へ追加・代替しない。
