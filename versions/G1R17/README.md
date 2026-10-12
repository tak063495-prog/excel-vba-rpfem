# G1R17 — P06A / P06N の粗メッシュ修復（DLL5改訂）

P06A・P06Nの粗メッシュで、下界荷重倍率がほぼ0に固定され、Fsが取得できない原因を修正しました。VBAとDLLの両方に修正を実装しています。

**追加試験（2026年10月13日）：同じG1R17で間隙水圧を含む26ケースのFs探索・適応解析を完了しました。Fs上下界取得26/26、精度目標達成10/26、主時間合計7時間46分7秒。[図入り実測報告・全結果・証拠ZIP](../../docs/benchmarks/G1R17_native26_20261013/README.md)を参照してください。以下の配布2冊は修復時の粗メッシュ設定を保持しています。**

|設定済みブック|Fs下界|Fs上界|今回の終端判定|
|---|---:|---:|---|
|[P06N・Davis等価](delivery/RPFEM_20261012_G1R17_P06N_Coarse_Checked.xlsm)|0.56014868|0.88211706|PARTIAL_CERTIFIED|
|[P06A・関連](delivery/RPFEM_20261012_G1R17_P06A_Coarse_Checked.xlsm)|0.73110643|1.02853738|PARTIAL_CERTIFIED|

両冊は同じコードで、入力だけが異なります。TARGET=200、実際245要素・147節点、q=8、適応OFF、浸透OFF、kh=0の**粗メッシュ確認用**です。上界で規準化した区間幅はP06N 36.50%、P06A 28.92%で、目標精度には未達です。この表は10月12日の修復検証時点の値です。後続のフル適応解析ではP06A/P06Nのq8上下界を取得しましたが、GAP目標は未達でした。単一の精密なFsとして扱わないでください。

`delivery/`のブックと[G1R17対応DLL](delivery/rpfem_native_x64.dll)を同じフォルダーに置き、Excelを終了してからDLLを交換します。原本は保管し、別名コピーで「要素定義」の「解析」を実行します。通常の利用にPythonやVBAプロジェクトへのアクセス許可は不要です。DLLがない場合、旧DLLの場合、32bit環境ではVBA経路になります。今回、DLL不在と旧DLLの実Excel復帰を確認しています。

- [原因・実装・検証報告](REPORT.md)
- [導入・手動インポート・復旧](IMPORT_ROLLBACK_JA.md)
- [試験台帳](TEST_RESULTS.csv) / [Python・Excelの再現手順](REPRODUCTION_JA.md)
- [VBA差分](changes.diff) / [DLLソース差分](native_changes.diff)
- [局所メッシュ比較図](figures/P06N_local_mesh.png)

DLL5の求解演算・MC・Davis変換・Fs定義・物理監査・許容差は維持しています。追加したのは、c=0自由頂点に接続する鈍角三角形の幾何判定と局所分割です。P06の番号や要素番号を製品コードに埋め込んでいません。

DLLはWindows x64、ABI `0x00170005`。G1R17のブックと組で使ってください。DLL6は作成していません。ライセンスは`native/licenses/`、配布ファイルのSHA-256は`SHA256SUMS.txt`に記載しています。
