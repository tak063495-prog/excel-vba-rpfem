param([string]$Src,[string]$Out,[string]$Tag)
$log="$Out\reg_$Tag.log"; function Log($m){ Add-Content $log ((Get-Date -f HH:mm:ss)+' '+$m) -Encoding UTF8 }
function Sheet($wb,$n){ foreach($sh in $wb.Sheets){ if($sh.Name -eq $n){ return $sh } } }
. (Join-Path $PSScriptRoot 'guard_excel.ps1'); try{ Stop-OrphanExcel }catch{ Log "$_"; exit 2 }
$dst="$Out\reg_$Tag.xlsm"; Copy-Item $Src $dst -Force
$xl=New-Object -ComObject Excel.Application; $xl.Visible=$false; $xl.DisplayAlerts=$false; $xl.AutomationSecurity=1; Set-RunMarker $Tag
try{
 $wb=$xl.Workbooks.Open($dst); $ws=Sheet $wb '設定'; $s=Sheet $wb '浸透入力'
 $set=@(); for($r=4;$r -le 16;$r++){ $set+=($ws.Cells.Item($r,4).Text+'='+$ws.Cells.Item($r,2).Text) }; Log ('settings ' + ($set -join ';') + ' coupling D4=' + $s.Range('D4').Text)
 try{ $xl.Run('RPX_SeepSetSilent',$true) }catch{}
 Log 'mesh'; $xl.Run('RPX_CreateMesh'); Log 'run'; $t=Get-Date; $xl.Run('RPX_Run'); Log ("run done " + [int]((Get-Date)-$t).TotalSeconds + "s")
 $r=Sheet $wb '解析結果'; Log ("A1=" + $r.Range('A1').Text + " lower=" + [string]$r.Range('B4').Value2 + " upper=" + [string]$r.Range('C4').Value2 + " elems=" + $r.Range('B10').Text + "/" + $r.Range('C10').Text + " H2=" + $r.Range('H2').Text)
 $wb.Save(); Log 'saved'
} catch { Log "ERR $_" } finally { try{$wb.Close($false)}catch{}; try{$xl.Quit()}catch{}; Clear-RunMarker }
