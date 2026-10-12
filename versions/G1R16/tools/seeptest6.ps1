param([string]$Book,[string]$Out,[string]$Tag)
. "$PSScriptRoot\common.ps1"; NewLog "$Out\seeptest_$Tag.log"
$dst="$Out\seeptest_$Tag.xlsm"; Copy-Item $Book $dst -Force
$wb=$null; $xl=StartXL
try{
 $wb=$xl.Workbooks.Open($dst)
 Log ('TestDam: ' + $xl.Run('RPX_SeepTestDam',40,20))
 foreach($r in 1e-8,0.5,1.0){ Log ('TestRelax: ' + $xl.Run('RPX_SeepTestRelax',[double]$r)) }
 try{ Log ('WT selftest: ' + $xl.Run('RPX_SeepWaterTableSelfTest')) }catch{ Log "WT selftest ERR $_" }
 # UI path with tiny relaxation: must report non-convergence and refuse coupling
 $s=Sheet $wb '浸透入力'; $row=0; for($i=1;$i -le 500;$i++){ if([string]$s.Cells.Item($i,1).Value2 -eq '緩和係数'){ $row=$i; break } }
 if($row -eq 0){ throw "relax label not found" }; Log "relax row=$row value=$($s.Cells.Item($row,3).Value2) D4=$($s.Range('D4').Value2) B3=$($s.Range('B3').Text)"
 Log "protected=$($s.ProtectContents)"; try{ $s.Unprotect() }catch{ Log "unprotect ERR $_" }; $xl.EnableEvents=$false; try{ SetC $s.Cells.Item($row,3) 1e-8 }catch{ Log "SetC ERR $_ sel=$($xl.ActiveSheet.Name) interactive=$($xl.Interactive) ready=$($xl.Ready)"; $s.Range("C$row").Formula="1E-08" }; $xl.EnableEvents=$true; Log "set relax=$($s.Cells.Item($row,3).Value2)"
 Log ('SeepRunSilent(relax=1e-8): [' + $xl.Run('RPX_SeepRunSilent') + ']')
 $r=Sheet $wb '浸透結果'; Log ("result K1(iters)=" + $r.Range('K1').Value2 + " K2(conv)=" + $r.Range('K2').Value2 + " W1=" + $r.Range('X1').Value2 + " W2=" + $r.Range('X2').Value2)
 $xl.Run('RPX_CreateMesh'); $xl.Run('RPX_Run'); $a=Sheet $wb '解析結果'; Log ("coupled run A1=" + $a.Range('A1').Text + " B4=" + $a.Range('B4').Text)
} catch { Log "ERR $_" } finally { EndXL $xl $wb }
