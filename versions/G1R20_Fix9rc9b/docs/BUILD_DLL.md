# Rebuilding rpfem_native_x64.dll (Winmax2, verified 2026-10-21)

Deployed DLL in every G1R19..G1R20_Fix9rc9a delivery: `rpfem_native_x64.dll`, sha256 `66da1a21be43cd2b3a3bfbaa9d13aad5fb1fd1960e1bf2507cc4e1cfa9ebe70b`, 103 exports (94 `rpx_*` ABI 0x00170005 + 9 `rpfem_accuracy_*` ABI 1).

## Source
- `legacy/` C11/C++17/MASM (MSVC `/fp:strict /MT`), `accuracy/` Rust staticlib (Clarabel 0.11.1 from crates.io, pinned by `Cargo.lock`, `--locked`), `accuracy_exports.def`, `build.ps1` / `build_legacy.cmd`.
- `build.ps1` checks `baseline_source_sha256.json` against `versions/G1R17/native`, `versions/G1R18/native`, `versions/G1R18/delivery` -> needs those paths from git (Winmax2 `rpfem_bench\repo` is a non-git copy without them).

## Toolchain (Winmax2)
- Rust 1.85.0 x86_64-pc-windows-msvc (LLVM 19.1.7). Use the original bootstrap `D:\Work\rpfem-rust-bootstrap\cargo` / `rustup` (the build.ps1 defaults). A user-level rustup 1.85.0 is also in `%USERPROFILE%\.cargo` but embeds different registry paths (not byte-equal).
- VS Build Tools 2022 17.14.33, MSVC 14.44.35207 (linker 14.44), Windows SDK 10.0.26100.0 / 10.0.22621.0.
- Only Windows PowerShell 5.1 is installed (no pwsh). `build.ps1` assumes pwsh 7 (cargo stderr becomes a terminating error; `GetRelativePath` missing). Use `build_ps51.ps1` (same steps; cargo via `cmd /c`, relative path via Substring).

## Commands (scratch, never inside a delivery folder)
```powershell
# box: cd fix9rc2_wt; git archive --format=zip -o src.zip HEAD versions/G1R19/native/unified versions/G1R17 versions/G1R18 validation/accuracy_20261013/python validation/accuracy_20261013/requirements.txt docs/benchmarks/G1R17_native26_20261013/data/inputs
cd C:\Users\link_\rpfem_bench\dll_build
Expand-Archive src.zip -DestinationPath repoN -Force
$env:Path="D:\Work\rpfem-rust-bootstrap\cargo\bin;$env:Path"
powershell -NoProfile -ExecutionPolicy Bypass -File repoN\versions\G1R19\native\unified\build_ps51.ps1   # ~2 min, runs 4 Rust unit tests
# headless regression (D:\Work\rpfem-venv has numpy 2.3.5, scipy 1.17.0, clarabel 0.11.1)
cd repoN; D:\Work\rpfem-venv\Scripts\python.exe -X utf8 versions\G1R19\native\unified\tests\verify_unified.py --dll <dll> --output <new folder>
```

## Equivalence result (unchanged source, 2026-10-21)
- Rebuilt sha `779aa8f7f6acf4b5716409594bc4fb499c23ed293b25c7e481dc444473d127f2` (not byte-equal to 66da1a21).
- `.text`, `.pdata`, `.reloc`, `.fptable` byte-identical; same size; 103/103 exports and dependents identical. 98 differing bytes only: /Brepro timestamp+repro hash, and MSVC anonymous-namespace RTTI names (`?A0x...`, derived from the source path; original was built under `D:\Work\excel-vba-rpfem`). Machine code identical -> results identical.
- verify_unified.py: both DLLs PASS (compile 120, meshfan 9, 17 FFI tests); result JSONs identical apart from hashes/times.
- Rust built with `%USERPROFILE%\.cargo` instead: sha 5dee5edd..., `.text` differs (8065 bytes) because embedded panic-path strings shift `.rdata`.

## Swapping a test DLL into a scratch workbook
VBA loads `rpfem_native_x64.dll` from the workbook folder. Copy a delivery folder (e.g. `versions\G1R20_Fix9rc9a\delivery`) to a new scratch folder, replace the DLL there, run the scratch xlsm. The build's "prepare" hash check expects 66da1a21; a modified DLL needs that expectation changed in the scratch harness only. Never overwrite delivery DLLs.
