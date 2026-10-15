# G1R19 単一 native DLL

配布時に必要な native DLL は **`rpfem_native_x64.dll` 1個**です。XLSM と同じフォルダーに置きます。精度回復用の別 DLL を読み込むラッパーではなく、従来の native 処理と Rust の処理を同じ PE ファイルへ静的リンクしています。

## ABI と構成

- 従来の `rpx_*` **94 exports**、`rpx_abi() = 0x00170005` を保持。
- `rpfem_accuracy_*` **9 exports**、`rpfem_accuracy_abi_version() = 1` を追加。元の SOCP/対角 QP worker と LSMR のデータ構造・所有権・キャンセル契約を保持。
- `legacy/` は G1R17 の C11/C++17/MASM ソースの保存コピー。浮動小数点の `/fp:strict` と MSVC `/MT` を維持。
- `accuracy/` は G1R18 の Rust ソースの保存コピー。crate type を `staticlib` に変更し、Rust **1.85.0**、Clarabel **0.11.1**、保存済み `Cargo.lock` を使用。Rust も `+crt-static` を使用。
- `accuracy_exports.def` が追加する9関数を明示。物理監査・根の判定条件は DLL 統合で変更していません。補助 QP の双対や solver status は元問題の証明にはなりません。

検証した PE の依存先は Windows 標準 DLL のみです。別の精度回復 DLL、VC runtime DLL は必要ありません。実際の一覧は [bin/dependencies.txt](bin/dependencies.txt)、公開関数は [bin/exports.txt](bin/exports.txt) に記録します。

## ビルド

**リポジトリ全体の checkout が前提**です。`baseline_source_sha256.json` は元 G1R17/G1R18 のソース・成果物262パスを参照し、ビルド前後で保全を確認します。配布 ZIP の DLL 単体ではこのビルド検証を行えません。配布 DLL の実行にはリポジトリは不要です。

Windows x64 の MSVC C++ build tools、Windows SDK、MASM、Rust 1.85.0 が必要です。リポジトリのルートで実行します。

```powershell
rustup toolchain install 1.85.0 --profile minimal
pwsh -File versions/G1R19/native/unified/build.ps1 `
  -CargoHome "$env:USERPROFILE/.cargo" `
  -RustupHome "$env:USERPROFILE/.rustup"
```

`build.ps1` は Rust staticlib のビルド・4 unit tests、C/C++/MASM のコンパイル、単一 DLL へのリンク、PE の export/依存先取得を行います。MSVC は `vswhere` で探し、Rust の追加システムライブラリーは `native-static-libs` の実出力を利用します。出力は `bin/rpfem_native_x64.dll` です。`bin/build_manifest.json` にソース・DLL の SHA256、toolchain、CRT、追加リンクライブラリーを保存します。

## Excel を起動しない回帰試験

既存の Python requirements を流用できます。大規模な P06 精度実験は実行しません。

```powershell
python -X utf8 -m pip install -r validation/accuracy_20261013/requirements.txt
python -X utf8 versions/G1R19/native/unified/tests/verify_unified.py `
  --dll versions/G1R19/native/unified/bin/rpfem_native_x64.dll `
  --output versions/G1R19/native/unified/validation/ci
```

`--output` には未使用のフォルダーを指定します。検証内容は次のとおりです。

- PE の AMD64、従来94＋追加9＝103 exports の完全一致、追加 DLL・VC runtime 依存の不在。
- 元 DLL と統合 DLL の LDL / double-double factor・backsolve **120ケースのビット一致**、double-double 演算・エラー応答。
- Python 独立再構成と native sparse compiler **120ケース**、実メッシュの fan/入力拒否/所有権 **9項目**、キャンセルした100 handle の解放。
- 元 q8 の8三角形 wedge・粘着材料、SOCP/対角 QP、入力コピー、CSR・非有限値拒否、worker キャンセル・結果取得・破棄、LSMR の SciPy 比較・元物理式への回復を含む **17 FFI tests**。

証拠は出力フォルダーの `summary.json`、`accuracy_ffi.json`、`accuracy_ffi.log`、`legacy/results/compile_python.json`、`legacy/results/native_meshfan.json` です。DLL の SHA256 を各試験の記録と照合します。Excel/VBA との結合、単一 DLL の実ロード、実務規模の時間測定は別の Excel 回帰試験で確認します。この headless 回帰だけで速度向上率や大規模 P06 再現を主張しません。

## ソースとライセンス

各成分のライセンスと通知を保存しています。従来 native 側の LGPL-2.1-or-later / Apache 通知は `legacy/licenses/` と各ソース、Rust 依存成分の通知は `accuracy/third_party/`、SciPy LSMR の BSD 通知は `accuracy/LICENSE_LSMR.txt` を参照してください。統合によりこれらの成分のライセンスは変更していません。保存ソース、固定依存、ビルド・リンク手順をリポジトリで提供し、DLL を再ビルドできる構成です。
