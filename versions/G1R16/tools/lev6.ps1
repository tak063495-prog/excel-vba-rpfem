param([string]$Book,[string]$Out,[string]$Tag)
. "$PSScriptRoot\common.ps1"; NewLog "$Out\lev_$Tag.log"
$dst="$Out\lev_$Tag.xlsm"; Copy-Item $Book $dst -Force
$wb=$null; $xl=StartXL
try{
 $wb=$xl.Workbooks.Open($dst); $ws=Sheet $wb '設定'; $s=Sheet $wb '浸透入力'
 $set=@(); for($r=4;$r -le 16;$r++){ $set+=($ws.Cells.Item($r,4).Text+'='+$ws.Cells.Item($r,2).Text) }; Log ('settings ' + ($set -join ';') + ' D4=' + $s.Range('D4').Text + ' B3=' + $s.Range('B3').Text)
 try{ $xl.Run('RPX_SeepSetSilent',$true) }catch{}
 Log ('SeepRunSilent: [' + $xl.Run('RPX_SeepRunSilent') + ']')
 $r=Sheet $wb '浸透結果'; Log ("seep iters=" + $r.Range('K1').Value2 + " conv=" + $r.Range('K2').Value2 + " qin=" + $r.Range('K4').Value2 + " qout=" + $r.Range('K5').Value2 + " dfix=" + $r.Range('X1').Value2 + " res=" + $r.Range('X2').Value2)
 Log 'mesh'; $xl.Run('RPX_CreateMesh'); Log 'run'; $t=Get-Date; $xl.Run('RPX_Run'); Log ("run done " + [int]((Get-Date)-$t).TotalSeconds + "s")
 $a=Sheet $wb '解析結果'; Log ("A1=" + $a.Range('A1').Text + " lower=" + [string]$a.Range('B4').Value2 + " upper=" + [string]$a.Range('C4').Value2 + " elems=" + $a.Range('B10').Text + "/" + $a.Range('C10').Text + " H2=" + $a.Range('H2').Text)
 $xl.Run('RPX_ShowDiagram'); $f=Sheet $wb '図'; $f.Activate(); $rg=$f.Range('A1:Z40'); $rg.CopyPicture(1,-4147); $co=$f.ChartObjects().Add(0,0,$rg.Width,$rg.Height); $co.Activate(); $co.Chart.Paste(); [void]$co.Chart.Export("$Out\lev_$Tag.png"); $co.Delete(); Log 'png'
 $wb.Save(); Log 'saved'
} catch { Log "ERR $_" } finally { EndXL $xl $wb }
