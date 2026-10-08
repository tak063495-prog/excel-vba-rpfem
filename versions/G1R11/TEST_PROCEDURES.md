# 検証手順

## 通常のP06N全解析（未実行）

1. 完成版G1R11の別コピーを開きます。解析結果の冒頭は全解析未実行の表示です。
2. P06N材料・形状があることと、設定B48=LEGACY_THEN_REFINE、B49=1000、B50=AUDITED_UPPER_LIMITを確認します。TARGET=1536、CYCLES=3、GAP=0.01、QUADRATURE=8、FS_TOL=0.0001、TEST_INJECTION=0は保持します。
3. 従来と同じ解析操作を行い、全解析終了までログと結果を保管します。
4. adapt_mesh_policyのstaged=True、従来のband経路、adapt_actionのlegacy_then_red_green、adapt_staged_endを確認します。追加段階は最大1回です。TARGETと追加上限の小さい方を超えていないことを確認します。
5. 下界でfs_upper_limit、必要ならfs_upper_limit_fallbackを確認します。上界を根拠に下界の不実行点を認証していないことを、fs_pointの実求解記録で確認します。
6. 最終audit、Fs根幅、adapt_end、outcome、run_endを確認します。PARTIAL_CERTIFIEDとGAP未達は全精度合格として扱いません。停止対象の例外が出た実行も合格にしません。
7. 下界・上界・GAP・要素数・求解数・全所要時間を記録します。比較値はG1R9の0.630355/0.823767、GAP26.60%、23,952.24秒と、G1R10の0.613174/0.711569、GAP14.85%、44,471.36秒です。ログの値だけを異なる実行間で合成しないでください。

1000要素の固定点試験を通ったことは、全適応経路と全Fs根を通った証拠ではありません。実行結果があるまではexcel_acceptance_g1r11.csvの全解析行をNOT_RUNのままにします。

## 25ケースと剛体（未実行）

validation_runtime/data/cases.jsonの26件から間隙水圧ケースD08Aを除く25ケースを、別コピーへ個別に設定します。P06Aは関連流れ則、P06Nはψ=0のDavis等価モデルとして区別します。単層、多層、c=0、小さい正のcを含め、各ケースの全Fs・監査・時間を記録します。実行したケースだけexcel_acceptance_g1r11.csvを更新します。今回の新しい探索は単層・剛体では無効です。剛体の通常SRM回帰は別のNOT_RUN行に残しています。

## 実施済み試験の再実行

PythonをRフォルダーで実行します。依存する保全原本はIMPORT_ROLLBACK.mdを参照してください。

```powershell
python -X utf8 tests/static_invariants.py
python -X utf8 tests/python_bracket_pretest.py
python -X utf8 python_staged_pretest.py
python -X utf8 native_geometry.py
python -X utf8 native_controlled.py
python -X utf8 native_brackets.py
python -X utf8 native_settings_guard.py
python -X utf8 native_small_roots.py
python -X utf8 native_fixed.py lower
python -X utf8 native_fixed.py upper
python -X utf8 audit_native.py
python -X utf8 verify_protection.py
```

依存関係のあるメッシュ生成→固定点求解→監査はこの順に行います。native_fixedの下界・上界は別Excelプロセスのテストコピーで実行しました。制御試験の代替値・代替監査は物理求解合格と混同しません。

release.pyは保全原本から成果物を再作成するスクリプトです。通常の検証には不要です。ユーザーが変更した成果物とSHA-256が異なる場合は停止するため、そのチェックを外して上書きしないでください。
