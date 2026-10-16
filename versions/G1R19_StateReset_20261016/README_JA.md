# 引数なしSubの明示Call候補

未公開・隔離overlay。Fsは検証済み `G1R19_2pct` のRPX_Fs、Conic/Solve/Seepは公式 `G1R19` を基準とし、各1行に `Call ` を追加した。元の `SubName:` はVBAでラベルとして解釈されるため、意図した呼び出しを明示する。

| モジュール | 明示した呼び出し |
|---|---|
| RPX_Conic | Farkas候補を検査する前のOriginalGuideClear |
| RPX_Fs | 新規Direct求解をCompileする前のResetPreparedState |
| RPX_Solve | native Krylovを準備する前のHContext |
| RPX_Seep | テスト補助関数SeepTestRelax内のSeepDefaults |

物理式、objective、q、原物理監査、Fs根の符号・許容幅、目標GAP、DLLの変更はない。現行コードには別のcontext/epoch/candidate検査が残り、この修正を「誤ったPASSの再現済み修正」とは扱わない。詳細な読み取り監査は `validation/G1R19_colon_review_20261016/REVIEW_JA.md` にある。

`src/vba` はUTF8、`import` は同じテキストのCP932、改行はCRLF。SOURCE_MANIFEST.jsonは各基準・候補・importの全byte SHAを記録する。元の4fix、WaterCommon、実行中の出力・ブックには書き込まない。

静的契約の確認:

```powershell
D:/Work/rpfem-venv/Scripts/python.exe -X utf8 -m unittest discover -s versions/G1R19_StateReset_20261016/tests -p test_contract.py -v
D:/Work/rpfem-venv/Scripts/python.exe -X utf8 versions/G1R19_StateReset_20261016/make_overlay_manifest.py --check
```

実Excel compile/reopenとmicroは未実行。`tests/MICROTEST_JA.md` の手順で、別の使い捨てブックを作る。`tests/make_probe_sources.py` はそのためのtest-only62+1モジュールを新規フォルダに作成する。sentinelと中断trapを含み、数値証拠を作る用途には使用しない。
