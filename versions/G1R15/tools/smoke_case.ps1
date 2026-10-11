param([string]$Src,[string]$Out,[string]$Tag,[int]$NoSeis=0,[int]$NoSeep=0,[int]$Vis=0,[double]$Kh=0.15,[int]$Target=300,[int]$Adapt=0,[string]$Dir='右(+x)',[int]$Seep=1,[string]$Kit='')
$log="$Out\iso_$Tag.log"; function Log($m){ Add-Content $log ((Get-Date -f HH:mm:ss)+' '+$m) -Encoding UTF8 }
function SetV($sh,$a,$v){ $r=$sh.Range($a); [void]$r.GetType().InvokeMember('Value2',[Reflection.BindingFlags]::SetProperty,$null,$r,@($v)) }
function Sheet($wb,$n){ foreach($sh in $wb.Sheets){ if($sh.Name -eq $n){ return $sh } } }
Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
$dst="$Out\iso_$Tag.xlsm"; Copy-Item $Src $dst -Force
$xl=New-Object -ComObject Excel.Application; $xl.Visible=[bool]$Vis; $xl.DisplayAlerts=$false; $xl.AutomationSecurity=1
try{
 $wb=$xl.Workbooks.Open($dst)
 if($Kit){ $vbp=$wb.VBProject; $o=$vbp.VBComponents.Item('RPX_Seismic'); $o.Name='RPX_Seismic_old'; $vbp.VBComponents.Remove($o); [void]$vbp.VBComponents.Import("$Kit\RPX_Seismic_cp932.bas"); Log 'seismic module replaced' }
 $cm=$wb.VBProject.VBComponents.Item('RPX_Control').CodeModule
 for($i=1;$i -le $cm.CountOfLines;$i++){ $t=$cm.Lines($i,1).Trim()
  if($NoSeis -and $t -eq 'Call RPX_SeismicOverlay'){ $cm.ReplaceLine($i,"    ' Call RPX_SeismicOverlay"); Log "disabled seis $i" }
  if($NoSeep -and $t -eq 'Call RPX_SeepOverlay'){ $cm.ReplaceLine($i,"    ' Call RPX_SeepOverlay"); Log "disabled seep $i" } }
 $ws=Sheet $wb '設定'; SetV $ws 'B5' $Target; SetV $ws 'B7' $Adapt; SetV $ws 'B15' $Kh; SetV $ws 'B16' $Dir
 if($Seep){ $s=Sheet $wb '浸透入力'; SetV $s 'D4' 1; Log ('SeepRunSilent: [' + $xl.Run('RPX_SeepRunSilent') + ']') }
 try{ $xl.Run('RPX_SeepSetSilent',$true) }catch{ Log 'no silent switch in this book' }
 Log "mesh kh=$Kh vis=$Vis noSeis=$NoSeis noSeep=$NoSeep"; $xl.Run('RPX_CreateMesh'); Log 'run'; $xl.Run('RPX_Run'); Log 'run done'
 $r=Sheet $wb '解析結果'; Log ("A1=" + $r.Range('A1').Text + " lower=" + $r.Range('B4').Text + " upper=" + $r.Range('C4').Text + " H2=" + $r.Range('H2').Text)
 $xl.Run('RPX_ShowDiagram'); Log 'diagram shown'
 $f=Sheet $wb '図'; $f.Activate(); $rg=$f.Range('A1:Z40'); $rg.CopyPicture(1,-4147); $co=$f.ChartObjects().Add(0,0,$rg.Width,$rg.Height); $co.Activate(); $co.Chart.Paste(); [void]$co.Chart.Export("$Out\iso_$Tag.png"); $co.Delete(); Log 'png'
 $wb.Save(); Log 'saved'
} catch { Log "ERR $_" } finally { try{ $wb.Close($false) }catch{}; try{ $xl.Quit() }catch{} }
