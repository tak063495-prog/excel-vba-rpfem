param([Parameter(Mandatory)][string]$BaseBook,[Parameter(Mandatory)][string]$SrcDir,[Parameter(Mandatory)][string]$OutDir)
# Builds baseline (unchanged) and patched copies, runs Issue15Probe in each, saves patched delivery book.
$ErrorActionPreference='Stop'
Add-Type -Namespace Issue15 -Name Win -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(System.IntPtr h, out uint p);'
function SourceText($path){ (([IO.File]::ReadAllText($path,[Text.Encoding]::UTF8) -split '\r?\n' | Where-Object {$_ -notlike 'Attribute VB_*'}) -join "`r`n").Trim() }
$running=@(Get-Process EXCEL -ErrorAction SilentlyContinue)
if($running.Count -gt 0){ throw "Another Excel is running (PIDs $($running.Id -join ',')); refusing" }
New-Item -ItemType Directory -Force $OutDir | Out-Null
$xl=New-Object -ComObject Excel.Application
[uint32]$excelPid=0; [void][Issue15.Win]::GetWindowThreadProcessId([IntPtr]$xl.Hwnd,[ref]$excelPid)
$xl.Visible=$false; $xl.DisplayAlerts=$false; $xl.EnableEvents=$false
$res=[ordered]@{excel_version=$xl.Version;excel_pid=$excelPid;base_sha256=(Get-FileHash $BaseBook).Hash.ToLower()}
$book=$null
function Compile($m){ $book.Activate(); $book.VBProject.VBComponents.Item($m).CodeModule.CodePane.Show(); $c=$xl.VBE.CommandBars.FindControl(1,578); if($c.Enabled){$c.Execute()}; if($c.Enabled){throw "compile failed"}; $xl.VBE.MainWindow.Visible=$false }
function Run($n,$a){ $xl.GetType().InvokeMember('Run',[Reflection.BindingFlags]::InvokeMethod,$null,$xl,@("'"+$book.Name+"'!"+$n)+$a) }
try{
  foreach($variant in 'baseline','patched'){
    $copy=Join-Path $OutDir "$variant-probe.xlsm"; Copy-Item -LiteralPath $BaseBook -Destination $copy -Force
    $book=$xl.Workbooks.Open($copy,0,$false)
    if($variant -eq 'patched'){ $cm=$book.VBProject.VBComponents.Item('RPX_SeepUI').CodeModule; $cm.DeleteLines(1,$cm.CountOfLines); $cm.AddFromString((SourceText (Join-Path $SrcDir 'RPX_SeepUI.bas'))) }
    $c=$book.VBProject.VBComponents.Add(1); $c.Name='Issue15Probe'; $c.CodeModule.AddFromString((SourceText (Join-Path $PSScriptRoot 'Issue15Probe.bas')))
    Compile 'Issue15Probe'
    $r=[ordered]@{}
    foreach($k in 'validation','unconverged','success'){ foreach($w in $false,$true){ $r["$k/"+$(if($w){'warm'}else{'cold'})]=Run 'Issue15Case' @($k,[bool]$w) } }
    $res[$variant]=$r
    $book.Close($false); $book=$null
  }
  # delivery book: base + patched RPX_SeepUI only, compile, save, reopen and compare
  $deliv=Join-Path $OutDir 'delivery.xlsm'; Copy-Item -LiteralPath $BaseBook -Destination $deliv -Force
  $book=$xl.Workbooks.Open($deliv,0,$false)
  $base=@{}; foreach($c in $book.VBProject.VBComponents){ $base[$c.Name]= if($c.CodeModule.CountOfLines){$c.CodeModule.Lines(1,$c.CodeModule.CountOfLines)}else{''} }
  $cm=$book.VBProject.VBComponents.Item('RPX_SeepUI').CodeModule; $cm.DeleteLines(1,$cm.CountOfLines); $cm.AddFromString((SourceText (Join-Path $SrcDir 'RPX_SeepUI.bas')))
  Compile 'RPX_SeepUI'; $book.Save(); $book.Close($false); $book=$null
  $book=$xl.Workbooks.Open($deliv,0,$true); $casing=@()
  foreach($c in $book.VBProject.VBComponents){ $a= if($c.CodeModule.CountOfLines){$c.CodeModule.Lines(1,$c.CodeModule.CountOfLines)}else{''}
    if($c.Name -eq 'RPX_SeepUI'){ if($a.Trim() -ine (SourceText (Join-Path $SrcDir 'RPX_SeepUI.bas'))){throw 'source mismatch'} }
    elseif($a -cne $base[$c.Name]){ if($a -ine $base[$c.Name]){throw "unintended change $($c.Name)"}; $casing+=$c.Name } }
  Compile 'RPX_SeepUI'; $book.Close($false); $book=$null
  $res.delivery=[ordered]@{compile='PASS';reopen='PASS';source_match=$true;casing_only=$casing;sha256=(Get-FileHash $deliv).Hash.ToLower()}
  $res.status='PASS'
} catch { $res.status='FAIL'; $res.error="$($_.Exception.Message)"; throw }
finally {
  $res | ConvertTo-Json -Depth 6 | Set-Content -Encoding utf8 (Join-Path $OutDir 'issue15_excel.json')
  try{ if($book){$book.Close($false)} }catch{}; try{$xl.Quit()}catch{}
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl); [GC]::Collect(); Start-Sleep 3
  if(Get-Process -Id $excelPid -ErrorAction SilentlyContinue){ Stop-Process -Id $excelPid -Force }
}
