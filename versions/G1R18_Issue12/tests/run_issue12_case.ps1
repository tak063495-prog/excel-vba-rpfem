param([Parameter(Mandatory)][string]$SourceBook,[Parameter(Mandatory)][string]$Name,[Parameter(Mandatory)][string]$RunDir,
      [string]$ModelJson='',[string]$Settings='',[Parameter(Mandatory)][string]$ProbePath)
# One analysis in one owned Excel. Source book is copied, never modified.
$ErrorActionPreference='Stop'
function ComValue($x){ if($x -is [string]){ return [string]$x }; return [double]$x }
function SetCell($ws,[int]$row,[int]$col,$x){ $addr=([string][char](64+$col))+$row; $rg=$ws.Range($addr); if($x -is [string]){ $rg.Value2=[string]$x } else { $rg.Value2=[double]$x } }
. (Join-Path $PSScriptRoot 'common_issue12.ps1')
New-Item -ItemType Directory -Force $RunDir | Out-Null
$stamp=Get-Date -Format 'yyyyMMdd_HHmm'
$book=Join-Path $RunDir ("{0}_{1}.xlsm" -f $Name,$stamp)
if(Test-Path -LiteralPath $book){throw "Refusing to overwrite $book"}
Copy-Item -LiteralPath $SourceBook -Destination $book
$marker=Join-Path $RunDir 'RUNNING.marker'
$r=[ordered]@{name=$Name;book=$book;source_sha256=(Get-FileHash $SourceBook).Hash.ToLower();settings=$Settings;model=$ModelJson}
$h=Start-OwnedExcel; $wb=$null
try{
    $xl=$h.xl; $r.excel_pid=$h.pid
    Set-Content -Encoding utf8 $marker "pid=$($h.pid);book=$book;started=$(Get-Date -Format o)"
    $xl.EnableEvents=$false
    $wb=$xl.Workbooks.Open($book,0,$false)
    $c=$wb.VBProject.VBComponents.Add(1); $c.Name='Issue12Probe'; $c.CodeModule.AddFromString((SourceText $ProbePath))
    Compile-Book $xl $wb 'Issue12Probe'
    if($ModelJson){
        $m=Get-Content -Raw -Encoding utf8 $ModelJson | ConvertFrom-Json
        $ws=$wb.Worksheets.Item('要素定義'); $ws.Range('A7:T1006').ClearContents()
        foreach($spec in @(@('A','points',3,1),@('E','regions',4,5),@('J','faces',11,10))){
            $row=7; foreach($rec in $m.($spec[1])){ for($i=0;$i -lt $spec[2];$i++){ SetCell $ws $row ($spec[3]+$i) $rec[$i] }; $row++ }
        }
        $ws=$wb.Worksheets.Item('材料データ'); $ws.Range('A5:Q1004').ClearContents()
        $row=5; foreach($rec in $m.materials){ for($i=0;$i -lt 5;$i++){ SetCell $ws $row (1+$i) $rec[$i] }; $row++ }
        $kv=@{}; $m.settings.PSObject.Properties | ForEach-Object { $kv[$_.Name]=$_.Value }
    } else { $kv=@{} }
    foreach($pair in ($Settings -split ',' | Where-Object {$_})){ $k,$v=$pair -split '=',2; $kv[$k]= if($v -match '^-?[0-9.]+(E-?[0-9]+)?$'){[double]$v}else{$v} }
    $sws=$wb.Worksheets.Item('設定'); $r.skipped_settings=@()
    foreach($k in $kv.Keys){
        $done=$false
        for($row=4;$row -le 1000 -and -not $done;$row++){ if([string]$sws.Cells.Item($row,4).Value2 -eq $k){ SetCell $sws $row 2 $kv[$k]; $done=$true } }
        if(-not $done){ if($k -like 'SEISMIC_*' -and [double]$kv['SEISMIC_KH'] -eq 0){ $r.skipped_settings+=@($k) } else { throw "Setting row not found: $k" } }
    }
    $wb.Worksheets.Item('浸透入力').Range('D4').Value2=0
    $xl.EnableEvents=$true
    $run={param($n,$a) $args2=@("'"+$wb.Name+"'!"+$n)+$a; $xl.GetType().InvokeMember('Run',[Reflection.BindingFlags]::InvokeMethod,$null,$xl,$args2)}
    [void](& $run 'RPX_NativeSetMode' @('auto')); $r.native=[string](& $run 'RPX_NativeInfo' @())
    if(-not $r.native.StartsWith('DLL_READY')){throw "DLL not ready: $($r.native)"}
    $wb.Save()
    $t=[Diagnostics.Stopwatch]::StartNew(); [void](& $run 'RPX_Run' @()); $r.seconds=$t.Elapsed.TotalSeconds
    $res=$wb.Worksheets.Item('解析結果')
    $r.terminal=[string]$res.Range('A1').Value2; $r.lower=$res.Range('B4').Value2; $r.upper=$res.Range('C4').Value2
    $r.state=[string](& $run 'Issue12State' @())
    $out=@(); for($i=1;$i -le 40;$i++){ $line=@(); for($j=1;$j -le 4;$j++){ $line+= [string]$res.Cells.Item($i,$j).Value2 }; if(($line -join '').Trim()){ $out+=($line -join ' | ') } }
    $r.result_sheet=$out
    $wb.Save()
    $log=Get-ChildItem (Join-Path $RunDir 'RPFEM_logs') -Filter '*analysis*.tsv' -ErrorAction SilentlyContinue | Sort-Object LastWriteTime | Select-Object -Last 1
    if($log){
        $L=Get-Content -LiteralPath $log.FullName -Encoding Unicode
        $r.log=$log.Name
        $r.cache_lookup_hits=@($L | Select-String 'lower_cache_lookup.*hit=1').Count
        $r.cache_lookup_miss=@($L | Select-String 'lower_cache_lookup.*hit=0').Count
        $r.cache_evidence_invalidations=@($L | Select-String 'lower_cache_lookup.*original_evidence').Count
        $r.cache_stores=@($L | Select-String "`tlower_cache_store`t").Count
        $r.cache_restores=@($L | Select-String "`tlower_cache_restore`t").Count
        $r.restore_original_q=@($L | Select-String 'lower_cache_restore' | ForEach-Object { if($_ -match 'original_q=(\d+)'){$matches[1]} } | Select-Object -Unique)
        $r.hsd_not_converged=@($L | Select-String 'HSD_NOT_CONVERGED').Count
        $r.original_lower_dual_guide=@($L | Select-String 'ORIGINAL_LOWER_DUAL_GUIDE').Count
        $r.prepared_drops=@($L | Select-String 'lower_prepared_drop').Count
        $r.final_eval=@($L | Select-String 'adapt_eval.*label=final' | ForEach-Object { $_.Line.Substring([Math]::Max(0,$_.Line.IndexOf('label='))) })
        $r.run_end=@($L | Select-String "`trun_end`t" | ForEach-Object { $_.Line.Substring([Math]::Max(0,$_.Line.IndexOf('status='))) })
    }
    $r.status='DONE'
} catch { $r.status='FAIL'; $r.error="$($_.Exception.Message)"; $r.error_position="$($_.InvocationInfo.PositionMessage)" }
finally {
    Stop-OwnedExcel $h $wb
    try{ $r.output_sha256=(Get-FileHash $book).Hash.ToLower() }catch{}
    $r | ConvertTo-Json -Depth 6 | Set-Content -Encoding utf8 (Join-Path $RunDir 'result.json')
    Remove-Item -LiteralPath $marker -ErrorAction SilentlyContinue
}
Write-Output ("{0}: {1} lower={2} upper={3} hits={4} hsdnc={5}" -f $Name,$r.status,$r.lower,$r.upper,$r.cache_lookup_hits,$r.hsd_not_converged)
