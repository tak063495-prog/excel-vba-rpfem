# G1R20 Fix1 再現テストと回帰の実行方法

Windows・64bit ExcelとPython環境を使います。数値参照ライブラリは
`validation/accuracy_20261013/requirements.txt`、Excelハーネスは加えて
`pywin32`、`psutil`、保存VBAの独立抽出は `oletools` を必要とします。
以下のコマンドは、GitHubから取得したリポジトリのルートで実行します。
過去版・入力・参照実装を併用するため、配布ZIPだけでは全検証を再実行できません。

各Excelテストは専用コピーだけを変更します。Excelをほかで開いていない状態で
1本ずつ実行し、既存の出力ディレクトリを再利用しないでください。

## 求解を伴わないテスト

```powershell
python -X utf8 versions/G1R20_Fix1/tests/run_portable_regression.py --output <新しい結果.json>
python -X utf8 versions/G1R20_Fix1/tests/UI_run_excel.py --execute --output <新しいUI結果フォルダー>
```

UIハーネスは配布済みG1R19/G1R20のコピーへ修正版63ソースを読み込みます。
Setupを呼ばず、G1R19の設定追加と物理入力・解析政策の保持を確認します。

```powershell
python -X utf8 versions/G1R20_Fix1/tests/run_water_excel.py
python -X utf8 versions/G1R20_Fix1/tests/run_water_negative_control.py
```

水理の正側テストは修正版63ソースを配布G1R20の専用コピーへ読み込み、
同秒更新・同じ入力スコープ内の更新・モード0/1/2/3・収束フラグFalse・保護シートと途中出力失敗を検査します。
private状態の観測とUI失効操作には、試験コピーだけへ追加する明示的な橋渡しを使います。
負の対照は元G1R20ブックで古い水頭の再利用が再現することを検査し、再現できたときにexit0を返します。
どちらも時刻やK9を捏造せず、実際の規定水位の生成・出力処理を使います。

portableは新15件と既存82件の計97件です。過去のP06複合証拠11件は、
元の絶対パスとRelease証拠ZIPの生データを必要とするため別集計です。
元の証拠を変更せず元作業場所で11件の合格も再確認しました。
新チェックアウトでの最初の失敗は `evidence/portable_initial_fail.json` に残します。

## 新規の数値回帰

```powershell
python -X utf8 versions/G1R20_Fix1/tests/run_cold_regression_excel.py --case D02A --precision-from-start --budget 600
```

入力からの新規解析で、過去の応力・速度・Fs・根を読み込みません。
試験コピーのSaveStateに読取り専用の観測を追加し、採用場と元の物理条件を再検査します。
G1R20改善設定・q8・GAP2%・両側の根の条件を使います。
監視は専用ExcelのPID・生成時刻・コマンドを照合し、取消要求と独立した時間上限監視を持ちます。

```powershell
python -X utf8 validation/G1R19_D02_route_audit_20261016/audit_cold_native.py --run <D02A結果フォルダー> --case D02A --lower 00004_lower --upper 00005_upper --output <新しい監査.json>
python -X utf8 versions/G1R20_Fix1/tests/run_water_native_fixed_excel.py --seconds 240 --solve-seconds 20
```

D08Aは最終配布ブックへ試験用観測モジュールだけを加え、新規の規定水位結果と
水位線適合メッシュで下側Fs=1・上側Fs=2を求解します。
独立Pythonの原q8再組立・物理監査・原CSR再検査まで自動実行します。
これは固定点回帰で、根探索やGAP2%完走のテストではありません。

Release別添の `EXECUTION_EVIDENCE.zip` は、各実行のソーススナップショット、
試験コピー、ログ、原CSRと場をリポジトリ相対パスで保持します。
`MANIFEST.json` に全メンバーのSHA256を記録し、失敗した試験ハーネスの先行実行も残します。

## 配布物の検査

```powershell
python -X utf8 versions/G1R20_Fix1/tools/check_release_sources.py --require-delivery
python -X utf8 versions/G1R20_Fix1/tools/audit_book_sources.py versions/G1R20_Fix1
```

63モジュールのUTF8/CP932、元G1R20との差分4モジュール、保存ブック内のソース、
単一DLL、配布ZIPの各メンバーを照合します。
元のG1R20の配布物や過去の数値証拠は書き換えません。
