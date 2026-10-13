param([string]$Root = (Split-Path $PSScriptRoot), [switch]$CompileOnly)
$ErrorActionPreference = 'Stop'
$Root = [IO.Path]::GetFullPath($Root)
$base = Join-Path (Split-Path $Root) 'G1R17'
$runDir = Join-Path $Root ('results/adapt_excel_' + [DateTime]::UtcNow.ToString('yyyyMMdd_HHmmss_fff'))
New-Item -ItemType Directory -Force $runDir | Out-Null
$sourceBook = Join-Path $base 'delivery/RPFEM_20261012_G1R17_P06N_Coarse_Checked.xlsm'
$testBook = Join-Path $runDir 'AdaptAccuracy_Test.xlsm'
Copy-Item -LiteralPath $sourceBook -Destination $testBook
Copy-Item -LiteralPath (Join-Path $base 'delivery/rpfem_native_x64.dll') -Destination (Join-Path $runDir 'rpfem_native_x64.dll')
Add-Type -Namespace AdaptAccuracy -Name Win -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(System.IntPtr h, out uint p);'
$prior = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
$xl = $null
$book = $null
$owned = $false
$record = [ordered]@{status='RUNNING';output=$runDir;prior_pids=$prior;runs=[ordered]@{}}
function InvokeMacro($name, $arguments) {
    $params = @("'" + $book.Name + "'!" + $name) + $arguments
    return $xl.GetType().InvokeMember('Run', [Reflection.BindingFlags]::InvokeMethod, $null, $xl, $params)
}
function ReplaceSource($name, $path) {
    $cm = $book.VBProject.VBComponents.Item($name).CodeModule
    $code = [IO.File]::ReadAllText($path, [Text.Encoding]::UTF8)
    $code = ($code -split '\r?\n' | Where-Object {$_ -notlike 'Attribute VB_*'}) -join "`r`n"
    $cm.DeleteLines(1, $cm.CountOfLines)
    $cm.AddFromString($code)
}
try {
    $xl = New-Object -ComObject Excel.Application
    [uint32]$ownedPid = 0
    [void][AdaptAccuracy.Win]::GetWindowThreadProcessId([IntPtr]$xl.Hwnd, [ref]$ownedPid)
    if ($prior -contains [int]$ownedPid) {throw 'Refusing to use an existing Excel process'}
    $owned = $true
    $record.pid = $ownedPid
    $record.excel_version = $xl.Version
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $xl.EnableEvents = $false
    $book = $xl.Workbooks.Open($testBook, 0, $false)
    foreach ($module in 'RPX_Adapt','RPX_Seep','RPX_SeepUI') {
        ReplaceSource $module (Join-Path $Root "src/vba/$module.bas")
    }
    $probe = $book.VBProject.VBComponents.Add(1)
    $probe.Name = 'AdaptAccuracyProbe'
    ReplaceSource 'AdaptAccuracyProbe' (Join-Path $PSScriptRoot 'AdaptAccuracyProbe.bas')
    $book.Activate()
    $book.VBProject.VBComponents.Item('RPX_Adapt').CodeModule.CodePane.Show()
    $ctl = $xl.VBE.CommandBars.FindControl(1, 578)
    if ($ctl.Enabled) {$ctl.Execute()}
    if ($ctl.Enabled) {throw 'VBA compilation failed'}
    $xl.VBE.MainWindow.Visible = $false
    $record.compile = 'PASS'
    Write-Output 'G1R18 adaptive candidate compile PASS'
    if (-not $CompileOnly) {
        foreach ($test in @(@('disabled',64,0), @('continued',256,2))) {
            $label = [string]$test[0]
            $trace = Join-Path $runDir ($label + '.tsv')
            $result = InvokeMacro 'AccuracyProbeRun' @([int]$test[1],[int]$test[2],$trace)
            if (-not $result.StartsWith('PASS;')) {throw "$label failed: $result"}
            $record.runs[$label] = $result
            Write-Output "$label $result"
        }
    }
    $book.Save()
    $record.status = 'PASS'
} catch {
    $record.status = 'FAIL'
    $record.error = $_.Exception.Message
    throw
} finally {
    $record | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $runDir 'summary.json') -Encoding UTF8
    if ($book) {$book.Close($false)}
    if ($xl -and $owned) {$xl.Quit()}
    if ($probe) {[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($probe)}
    if ($book) {[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)}
    if ($xl) {[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl)}
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
