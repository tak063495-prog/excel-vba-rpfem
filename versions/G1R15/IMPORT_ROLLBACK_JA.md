# G1R15 導入・復旧

## 導入（推奨：tools/install_g1r15.ps1）
前提：Excelのトラストセンターで「VBAプロジェクト オブジェクト モデルへのアクセスを信頼する」をON。Excelは1つだけ起動。

```
powershell -ExecutionPolicy Bypass -File tools\install_g1r15.ps1 -Src <G1R14ブック> -Dst <新しい名前.xlsm> -Kit import -Log install.log
```
処理：`RPX_Seep`・`RPX_SeepUI`・`RPX_SeepWT`・`RPX_Water`・`RPX_Seismic` を差し替え／追加（旧モジュールは改名してから削除）、フック4行を冪等に挿入、`RPX_SeepSetupSheet`（浸透入力の移行）と`RPX_SeismicSetup`（設定15・16行の追加）を実行して保存。元ブックは変更しない（Dstが既存なら停止）。

## 手動導入
1. 上記5モジュールを import/ からインポート（同名がある場合は先に削除）。
2. `RPX_Mesh.RPX_AssignRegions` の自重行
   `If RGravityMode = "reference" Then RBody1(count, 1) = ... Else RBody0(count, 1) = ...`
   の直後に `            RPX_SeismicBody count`
3. `RPX_Control.RPX_DrawCore` の `Call RPX_SeepOverlay` の直後に `    Call RPX_SeismicOverlay`
4. G1R14未導入のブックでは G1R14 のフック2行（`Call RPX_WaterApply`、`Call RPX_SeepOverlay`）も必要。
5. デバッグ→コンパイル後、マクロ `RPX_SeepSetupSheet`、`RPX_SeismicSetup` を実行して保存。

## 復旧
- 地震を使わない：設定B15を0にすればG1R14と同じ体積力（S1の数値一致を確認済み）。
- 完全に戻す：G1R14のブックを使う（G1R15作成時に元ブックは変更していない）。
