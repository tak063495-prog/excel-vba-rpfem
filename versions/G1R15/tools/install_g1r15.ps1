param([string]$Src,[string]$Dst,[string]$Kit,[string]$Log)
function Log($m){ Add-Content $Log ((Get-Date -f HH:mm:ss)+' '+$m) -Encoding UTF8 }
if(Get-Process EXCEL -ErrorAction SilentlyContinue){ throw 'Excel is running (maybe mid-analysis). Close it yourself and retry; this script never kills Excel.' }
if(Test-Path $Dst){ throw "exists $Dst" }; Copy-Item $Src $Dst
$xl=New-Object -ComObject Excel.Application; $xl.Visible=$false; $xl.DisplayAlerts=$false; $xl.AutomationSecurity=1
try{
 $wb=$xl.Workbooks.Open($Dst); $vbp=$wb.VBProject
 $names=@(); foreach($c in $vbp.VBComponents){ $names+=$c.Name }
 foreach($m in 'RPX_Seep','RPX_SeepUI','RPX_SeepWT','RPX_Water','RPX_Seismic'){ if($names -contains $m){ $o=$vbp.VBComponents.Item($m); $o.Name=$m+'_old'; $vbp.VBComponents.Remove($o) }; [void]$vbp.VBComponents.Import("$Kit\${m}_cp932.bas") }
 foreach($h in @(@('RPX_Mesh','If RPX_RepairLoadSteps() Then Call RPX_AssignRegions','    Call RPX_WaterApply'),@('RPX_Mesh','If RGravityMode = "reference" Then RBody1(count, 1)*','            RPX_SeismicBody count'),@('RPX_Control','If RPX_ResultCurrent Then Call RPX_DrawVelocity','    Call RPX_SeepOverlay'),@('RPX_Control','Call RPX_SeepOverlay','    Call RPX_SeismicOverlay'))){
  $cm=$vbp.VBComponents.Item($h[0]).CodeModule; $hit=0;$done=$false
  for($i=1;$i -le $cm.CountOfLines;$i++){ $t=$cm.Lines($i,1).Trim(); if($t -like $h[1]){$hit=$i}; if($t -eq $h[2].Trim()){$done=$true} }
  if(-not $done){ if($hit -eq 0){ throw "anchor missing $($h[0]) $($h[1])" }; $cm.InsertLines($hit+1,$h[2]) }; Log "hook $($h[0]) [$($h[2].Trim())] at $hit already=$done" }
 $n2=@(); foreach($c in $vbp.VBComponents){ $n2+=$c.Name }; Log ('modules: ' + (($n2 | ?{$_ -like 'RPX_Seep*' -or $_ -like 'RPX_Water*' -or $_ -like 'RPX_Seis*'}) -join ','))
 $xl.Run('RPX_SeepSetSilent',$true); $xl.Run('RPX_SeepSetupSheet'); $xl.Run('RPX_SeismicSetup'); Log ('setup msgs: ' + $xl.Run('RPX_SeepGetLastMsg')); $xl.Run('RPX_SeepSetSilent',$false)
 Log ('compile probe: ' + $xl.Run('RPX_SeismicNote') + ' / ' + $xl.Run('RPX_SeepGetLastMsg'))
 foreach($sh in $wb.Sheets){ if($sh.Name -eq '設定'){ Log ('B15=' + $sh.Range('B15').Text + ' D15=' + $sh.Range('D15').Text + ' B16=' + $sh.Range('B16').Text + ' D16=' + $sh.Range('D16').Text) } }
 $wb.Save(); Log "saved $Dst"
} catch { Log "ERR $_" } finally { try{$wb.Close($false)}catch{}; $xl.Quit() }
