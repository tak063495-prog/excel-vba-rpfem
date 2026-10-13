param(
    [string]$CargoHome='D:\Work\rpfem-rust-bootstrap\cargo',
    [string]$RustupHome='D:\Work\rpfem-rust-bootstrap\rustup',
    [switch]$SkipTests
)
$ErrorActionPreference='Stop'
$taskOriginalCargo=$env:CARGO_HOME
$taskOriginalRustup=$env:RUSTUP_HOME
$taskOriginalDirectory=Get-Location
try {
    $env:CARGO_HOME=$CargoHome
    $env:RUSTUP_HOME=$RustupHome
    $taskCargo=Join-Path $CargoHome 'bin\cargo.exe'
    if (-not (Test-Path -LiteralPath $taskCargo -PathType Leaf)) { throw 'Specify the isolated cargo runtime with -CargoHome.' }
    if (-not [Environment]::Is64BitProcess) { throw 'A Windows x64 build host is required.' }
    Set-Location -LiteralPath $PSScriptRoot
    # Rust discovers Visual Studio 2022 MSVC and Windows SDK through the local
    # installed-toolchain registry; no global PATH or HOME changes are made.
    $taskTarget='x86_64-pc-windows-msvc'
    & $taskCargo build --release --locked --target $taskTarget
    if ($LASTEXITCODE -ne 0) { throw 'Cargo build failed.' }
    if (-not $SkipTests) {
        & $taskCargo test --release --locked --target $taskTarget
        if ($LASTEXITCODE -ne 0) { throw 'Cargo tests failed.' }
    }
    $taskDestination=Join-Path $PSScriptRoot 'bin'
    New-Item -ItemType Directory -Path $taskDestination -Force | Out-Null
    $taskBuilt=Join-Path $PSScriptRoot "target\$taskTarget\release\rpfem_accuracy_x64.dll"
    $taskDll=Join-Path $taskDestination 'rpfem_accuracy_x64.dll'
    Copy-Item -LiteralPath $taskBuilt -Destination $taskDll -Force
    if ((Get-FileHash -LiteralPath $taskBuilt -Algorithm SHA256).Hash -ne (Get-FileHash -LiteralPath $taskDll -Algorithm SHA256).Hash) {
        throw 'Copied DLL SHA256 differs from the build artifact.'
    }
    $taskSourceHashes=[ordered]@{}
    $taskSources=@('Cargo.toml','Cargo.lock','.cargo\config.toml','rpfem_accuracy.h','build.ps1')
    $taskSources+=@(Get-ChildItem -LiteralPath (Join-Path $PSScriptRoot 'src') -Filter '*.rs' -File | ForEach-Object { 'src\'+$_.Name })
    foreach ($taskSource in $taskSources) {
        $taskSourceHashes[$taskSource]=(Get-FileHash -LiteralPath (Join-Path $PSScriptRoot $taskSource) -Algorithm SHA256).Hash.ToLowerInvariant()
    }
    $taskRustc=Join-Path $CargoHome 'bin\rustc.exe'
    $taskManifest=[ordered]@{
        target=$taskTarget
        clarabel='0.11.1'
        rustc=((& $taskRustc --version --verbose) -join "`n")
        cargo=((& $taskCargo --version) -join "`n")
        crt='static; PE /Brepro enabled'
        source_sha256=$taskSourceHashes
        dll_sha256=(Get-FileHash -LiteralPath $taskDll -Algorithm SHA256).Hash.ToLowerInvariant()
        rust_tests_skipped=[bool]$SkipTests
    }
    $taskManifest | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $taskDestination 'build_manifest.json') -Encoding UTF8
    Get-FileHash -LiteralPath $taskDll -Algorithm SHA256
} finally {
    Set-Location -LiteralPath $taskOriginalDirectory
    $env:CARGO_HOME=$taskOriginalCargo
    $env:RUSTUP_HOME=$taskOriginalRustup
}
