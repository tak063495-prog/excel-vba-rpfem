# G1R14：インポートと復旧

配布する`delivery/`のXLSMには実装済みコードと「浸透入力」シートが入っている。通常はインポート作業は不要。G1R13ブックは保存したまま、修正版を別名で開いて使う。

## G1R13ブックへ手動で反映する場合

1. G1R13ブック（D03N/D07A設定済み版または同一系列）を別名で保存する。
2. 新規標準モジュール`RPX_Seep`、`RPX_SeepUI`、`RPX_Water`を`import/`のCP932・CRLF版からインポートする。同名の旧版が残っている場合は解放後に一度保存して開き直してからインポートする（`RPX_Seep1`等の別名取り込みを防ぐ）。
3. 既存標準モジュールへ次の2行を追加する。`import/`の`RPX_Mesh`・`RPX_Control`で置換してもよい（内容はG1R13＋各1行）。
   - `RPX_Mesh.RPX_AssignRegions`：`If RPX_RepairLoadSteps() Then Call RPX_AssignRegions`の直後に`Call RPX_WaterApply`
   - `RPX_Control.RPX_DrawCore`：`If RPX_ResultCurrent Then Call RPX_DrawVelocity`の直後に`Call RPX_SeepOverlay`
4. 「デバッグ→VBAProjectのコンパイル」を実行し、`RPX_SeepSetupSheet`で浸透入力シートを作成する。イミディエイトで`?RPX_SeepTestDam(40,20)`と`?RPX_SeepOverlaySelfTest`を確認する。

G1R12ではE1の位置が`RPX_LoadPhase = "raw"`の直後だった。G1R13でこの行は`RPX_Topology`の前へ移動したため、同じ位置へ入れると旧メッシュの勾配で体積力を作る誤りになる。必ず荷重段差修復の行の後に入れる。

`tools/install_into_workbook.ps1`は版を判別して同じ2行を挿入する。実行にはVBAプロジェクトへの信頼アクセスが必要で、スクリプトはその設定やレジストリを変更しない。`RPX_CouplingInstall`は同じブック内からApplication.Runで呼んでも実行中のプロジェクト自身は書き換えられないため、通常は手作業かPowerShellを使う。

## 復旧

G1R13ブックを開くことで変更前へ戻せる。コードだけを戻す場合は、`RPX_Mesh`・`RPX_Control`を`before/vba/`へ戻し、新規3モジュールと「浸透入力」「浸透結果」シートを削除してコンパイルする。連成ON（N19=1）で得た結果を、連成OFFの結果として扱わない。
