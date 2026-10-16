# D02Aログ受領後の検索状況

最新のGitHub recursive main treeを再取得し、PR #20および main `303922050bea7ddaaf6f26fe560c1a4d10c76121` にD02A正本が公開されたことを確認した。ログ13ファイルの原バイト、book/DLL/input SHA、実設定・phase/callは [D02_LOG_STATS.json](D02_LOG_STATS.json) と [比較報告](PYTHON_VBA_COMPARISON_JA.md) に記録した。

以下は新着PR公開前の検索記録として保持する。現在の未確認状態を意味しない。

# D02Aタイムアウト報告の正本検索状況

確認時刻UTC: 2026-10-16T11:14:36.513862+00:00。GitHub RESTをgh経由で読み取り、原JSONを同じフォルダーに保全した。

- 全issues/PR本文 19件、更新順コメント 9件にD02の記述はなかった。
- mainの `validation/G1R19_preset2pct_20261016/D02A` 履歴0件、D02名の公開branch0件。
- 現時点PR #19のheadは `b70e93ba3d57cd802b8c28928c08e6867c776a9e` で、追加ログはD01Aのみ。D02Aの実タイムアウトの停止段階・実設定・公式book/DLLのSHAは未確認。D01AのSHAをD02Aの実行証拠へ転記しない。
- canonical D02A入力は `docs/benchmarks/G1R17_native26_20261013/data/inputs/D02A.json`、生バイトSHA `97c4f88a81575120452a427fdc80621a630a0abf4f09d5da119b3a64f6512a5b`、canonical JSON SHA `3911bed529d59055e7a1c1d928fba7adf667b7d00118f2331e01a643de53bf99`。

D02Aはc=0/phi=psi45°/gamma20、外形(0,0),(70,0),(70,10),(50,10),(20,30),(0,30)、面積1400、関連流れ/乾燥/外載なし。Pythonの幾何候補・成功結果とVBAの実タイムアウト原因の比較は、正本ログを入手して確定する。公開ログ未確認を、タイムアウト報告自体の否定には使わない。
