param([string]$Src,[string]$Dst,[string]$Kit,[string]$LogFile)
. "$PSScriptRoot\common.ps1"; NewLog $LogFile
if(Test-Path $Dst){ Log "exists $Dst"; exit 3 }; Copy-Item $Src $Dst
$wb=$null; $xl=StartXL
try{
 $wb=$xl.Workbooks.Open($Dst); $vbp=$wb.VBProject
 foreach($m in 'RPX_Seep','RPX_SeepUI','RPX_SeepWT'){ $o=$vbp.VBComponents.Item($m); $o.Name=$m+'_old'; $vbp.VBComponents.Remove($o); [void]$vbp.VBComponents.Import("$Kit\${m}_cp932.bas") }
 $n2=@(); foreach($c in $vbp.VBComponents){ $n2+=$c.Name }; Log ('modules: ' + (($n2 | ?{$_ -like 'RPX_Seep*'}) -join ','))
 Log ('compile=' + (Compile $xl $wb))
 Log ('probe relax=0.5: ' + $xl.Run('RPX_SeepTestRelax',0.5))
 $wb.Save(); Log "saved $Dst"
} catch { Log "ERR $_" } finally { EndXL $xl $wb }
