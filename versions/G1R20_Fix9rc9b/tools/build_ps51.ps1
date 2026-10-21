param(
    [string]$CargoHome='D:\Work\rpfem-rust-bootstrap\cargo',
    [string]$RustupHome='D:\Work\rpfem-rust-bootstrap\rustup',
    [switch]$SkipTests
)
$ErrorActionPreference='Stop'
$taskPreviousCargo=$env:CARGO_HOME
$taskPreviousRustup=$env:RUSTUP_HOME
$taskPreviousDirectory=Get-Location
$taskRepo=(Resolve-Path (Join-Path $PSScriptRoot '..\..\..\..')).Path
try {
    if (-not [Environment]::Is64BitProcess) { throw 'Windows x64 host is required.' }
    $env:CARGO_HOME=$CargoHome
    $env:RUSTUP_HOME=$RustupHome
    $taskCargo=Join-Path $CargoHome 'bin\cargo.exe'
    if (-not (Test-Path -LiteralPath $taskCargo)) { throw 'Cargo runtime is unavailable.' }
    $taskBaseline=Get-Content -Raw -LiteralPath (Join-Path $PSScriptRoot 'baseline_source_sha256.json') | ConvertFrom-Json
    foreach ($taskItem in $taskBaseline.PSObject.Properties) {
        $taskSource=Join-Path $taskRepo $taskItem.Name
        if ((Get-FileHash -LiteralPath $taskSource -Algorithm SHA256).Hash.ToLowerInvariant() -ne $taskItem.Value) { throw "Original artifact changed: $($taskItem.Name)" }
    }
    New-Item -ItemType Directory -Force -Path (Join-Path $PSScriptRoot 'obj\legacy'),(Join-Path $PSScriptRoot 'bin') | Out-Null
    Set-Location -LiteralPath (Join-Path $PSScriptRoot 'accuracy')
    $taskRustOutput=@(& cmd.exe /d /c "`"$taskCargo`" rustc --release --locked --target x86_64-pc-windows-msvc -- --print native-static-libs 2>&1")
    $taskRustExit=$LASTEXITCODE
    $taskRustOutput | Set-Content -LiteralPath (Join-Path $PSScriptRoot 'obj\rust_build.log') -Encoding UTF8
    $taskRustOutput | ForEach-Object { Write-Output $_ }
    if ($taskRustExit -ne 0) { throw 'Rust static library build failed.' }
    $taskLibraries=@($taskRustOutput | ForEach-Object { if ("$_" -match 'native-static-libs:\s*(.*)$') { $Matches[1] } })
    if ($taskLibraries.Count -ne 1 -or $taskLibraries[0] -notmatch 'kernel32.lib') { throw 'Rust native static library list is unavailable.' }
    $taskLibraries[0] | Set-Content -LiteralPath (Join-Path $PSScriptRoot 'obj\rust_native_static_libs.rsp') -Encoding ASCII
    if (-not $SkipTests) {
        & $taskCargo test --release --locked --target x86_64-pc-windows-msvc
        if ($LASTEXITCODE -ne 0) { throw 'Rust unit tests failed.' }
    }
    $taskVswhere=Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
    if (-not (Test-Path -LiteralPath $taskVswhere)) { throw 'Visual Studio build tools are unavailable.' }
    $taskVs=@(& $taskVswhere -latest -products '*' -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath)
    if ($taskVs.Count -ne 1) { throw 'Select one MSVC x64 build installation.' }
    $taskVcvars=Join-Path $taskVs[0] 'VC\Auxiliary\Build\vcvars64.bat'
    & (Join-Path $PSScriptRoot 'build_legacy.cmd') $taskVcvars
    if ($LASTEXITCODE -ne 0) { throw 'Unified C/C++/MASM link failed.' }
    foreach ($taskItem in $taskBaseline.PSObject.Properties) {
        if ((Get-FileHash -LiteralPath (Join-Path $taskRepo $taskItem.Name) -Algorithm SHA256).Hash.ToLowerInvariant() -ne $taskItem.Value) { throw "Original artifact changed after build: $($taskItem.Name)" }
    }
    $taskSources=[ordered]@{}
    foreach ($taskFolder in @('legacy','accuracy\src','accuracy\.cargo','accuracy\third_party')) {
        Get-ChildItem -LiteralPath (Join-Path $PSScriptRoot $taskFolder) -Recurse -File | ForEach-Object {
            $taskRelative=$_.FullName.Substring($PSScriptRoot.Length).TrimStart([char]92)
            $taskSources[$taskRelative]=(Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash.ToLowerInvariant()
        }
    }
    foreach ($taskName in @('accuracy\Cargo.toml','accuracy\Cargo.lock','accuracy\rust-toolchain.toml','accuracy\rpfem_accuracy.h','accuracy\LICENSE_LSMR.txt','build.ps1','build_legacy.cmd','accuracy_exports.def','baseline_source_sha256.json')) {
        $taskSources[$taskName]=(Get-FileHash -LiteralPath (Join-Path $PSScriptRoot $taskName) -Algorithm SHA256).Hash.ToLowerInvariant()
    }
    $taskDll=Join-Path $PSScriptRoot 'bin\rpfem_native_x64.dll'
    $taskRustc=Join-Path $CargoHome 'bin\rustc.exe'
    $taskManifest=[ordered]@{
        artifact='rpfem_native_x64.dll';legacy_abi='0x00170005';accuracy_abi=1
        target='x86_64-pc-windows-msvc';clarabel='0.11.1';crt='MSVC /MT and Rust +crt-static'
        rustc=((& $taskRustc --version --verbose) -join "`n")
        native_static_libraries=$taskLibraries[0];rust_tests_skipped=[bool]$SkipTests
        original_artifacts_unchanged=$true;source_sha256=$taskSources
        dll_sha256=(Get-FileHash -LiteralPath $taskDll -Algorithm SHA256).Hash.ToLowerInvariant()
    }
    $taskManifest | ConvertTo-Json -Depth 7 | Set-Content -LiteralPath (Join-Path $PSScriptRoot 'bin\build_manifest.json') -Encoding UTF8
    Get-FileHash -LiteralPath $taskDll -Algorithm SHA256
} finally {
    Set-Location -LiteralPath $taskPreviousDirectory
    $env:CARGO_HOME=$taskPreviousCargo
    $env:RUSTUP_HOME=$taskPreviousRustup
}
