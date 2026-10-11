# G1R15 DLL2・堤防例題

前回DLL **443.89秒 → 400.16秒**。
追加で **9.85%短縮**。元VBAの既存622.12秒からは **35.68%短縮・1.555倍**。
Fs下界1.224448、上界1.234308、監査603件PASS、同じ53求解。

## 使い方

1. ZIPを展開する。
2. `delivery/` のXLSMとこのZIPのDLLを同じフォルダーへ置く。
3. 現在の許可済みExcel設定で開き、通常どおり解析する。

DLLなしの場合はVBAへ自動切替する。64bit Excelで検証済み。
前回版と同じDLL名のため、前回の解析用Excelを閉じてから新しい版を開く。
ロード済みのDLLはそのExcel終了まで置き換えず、フォルダーを分けて利用する。
旧DLLと新XLSMを混ぜた場合は、export/ABIチェックによりVBAへ戻る。

条件：堤防例題、浸透連成ON、kh=0.15右、TARGET=600、メッシュ最適化ON。
比較の範囲と試験結果は [REPORT.md](REPORT.md)。導入・復旧は [IMPORT_ROLLBACK_JA.md](IMPORT_ROLLBACK_JA.md)。
