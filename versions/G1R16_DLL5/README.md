# G1R16 DLL5

通常内点法とHSDを含む追加最適化版です。堤防例題はDLL4 272.97秒からDLL5 142.17秒へ、47.92%短縮しました。Fs・入力・メッシュは一致し、全53求解/603監査の記録も一致しました。同一PC各1回の実測です。

`delivery/`のXLSMとDLLを組で同じフォルダーに置きます。DLL交換前にExcelを終了し、原本を残して別名で保存してください。DLLがない場合・非対応版の場合は従来VBAで計算します。マクロの利用は組織の規則に従ってください。通常の解析にPythonやVBAプロジェクトアクセスの許可は不要です。

配布XLSMはTARGET=600、q=8、適応ON、kh=0.15右、浸透ONの堤防例題です。「要素定義」の「解析」で再計算できます。Fs下界1.2244477290215、上界1.2343080287831916の結果を保存しています。

- [検証報告・速度比較](REPORT.md)
- [実装範囲・コードの対応](IMPLEMENTATION_JA.md)
- [導入・復旧](IMPORT_ROLLBACK_JA.md)
- [試験結果](TEST_RESULTS.csv) / [再現手順](tools/REPRODUCTION_JA.md)
- [VBA差分](changes.diff) / [ネイティブ差分](native_changes.diff)
- [編集用ソース](src/vba/) / [VBEインポート用](import/) / [DLLソース・ビルド](native/)

P06A/P06Nは固定Fs比較と限定した粗メッシュ比較です。フル適応解析、極小Fs探索、26ケース全再試験は未実施です。粗メッシュP06Nの基準版と同じFs未取得も記録しています。`PARTIAL_CERTIFIED`と目標精度達成を区別してください。

ABI `0x00160005`、Windows x64。ライセンスは`native/licenses/`、ファイルのハッシュは`SHA256SUMS.txt`を確認してください。
