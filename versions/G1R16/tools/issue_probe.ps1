param([string]$Book,[string]$Out,[string]$Tag)
. "$PSScriptRoot\common.ps1"; NewLog "$Out\issueprobe_$Tag.log"
$P="$PSScriptRoot\probe"; $dst="$Out\issueprobe_$Tag.xlsm"; Copy-Item $Book $dst -Force
$fail=0; function Check($ok,$m){ if($ok){ Log "PASS $m" } else { Log "FAIL $m"; $script:fail++ } }
function Strip($t){ ($t -split "`r?`n" | ?{ $_ -notlike 'Attribute VB_*' }) -join "`r`n" }
$wb=$null; $xl=StartXL
try{
 $wb=$xl.Workbooks.Open($dst); $vbp=$wb.VBProject
 $ws=Sheet $wb '設定'; $row=0; for($i=4;$i -le 200;$i++){ if([string]$ws.Cells.Item($i,4).Value2 -eq 'TEST_INJECTION'){ $row=$i; break } }
 if($row -eq 0){ for($i=4;$i -le 400;$i++){ if([string]$ws.Cells.Item($i,4).Value2 -eq ''){ $row=$i; break } }; SetC $ws.Cells.Item($row,4) 'TEST_INJECTION'; Log "TEST_INJECTION row added at $row" }
 SetC $ws.Cells.Item($row,2) 1
 $vbp.VBComponents.Item('RPX_Adapt').CodeModule.AddFromString((Strip (Get-Content "$P\AdaptIssueProbe.txt" -Raw -Encoding UTF8)))
 foreach($n in 'IssueSolverProbe','MeshProbe','SmallRootProbe'){ $c=$vbp.VBComponents.Add(1); $c.Name=$n; $c.CodeModule.AddFromString((Strip (Get-Content "$P\$n.bas" -Raw -Encoding UTF8))) }
 $cm=$vbp.VBComponents.Item('RPX_Solve').CodeModule; $t=$cm.Lines(1,$cm.CountOfLines)
 $m1="    RPX_DiagEnd dpLDL`r`n    If Not factorOK Then"
 if(([regex]::Matches($t,[regex]::Escape($m1))).Count -ne 1){ throw 'marker1' }
 $t=$t.Replace($m1,"    RPX_DiagEnd dpLDL`r`n    If IssueForceFactor Then IssueForceFactor = False: factorOK = False: RPX_PivotFailure = 123`r`n    If Not factorOK Then")
 $m2='Public Sub RPX_OptimizeHSD(Optional ByVal prepared As Boolean = False, Optional ByVal auditMode As String = "")'
 $pos=$t.IndexOf($m2); if($pos -lt 0){ throw 'marker2' }
 $tail=$t.Substring($pos); $k=$tail.IndexOf('    On Error GoTo Failed'); $tail=$tail.Substring(0,$k)+"    On Error GoTo Failed`r`n    IssueHSDCalls = IssueHSDCalls + 1`r`n    RPX_TestHook ""issue_hsd"""+$tail.Substring($k+'    On Error GoTo Failed'.Length)
 $t=$t.Substring(0,$pos)+$tail; $cm.DeleteLines(1,$cm.CountOfLines); $cm.AddFromString($t)
 $ok=Compile $xl $wb; Check $ok 'compile'; if(-not $ok){ throw 'compile failed' }
 $q="'"+$wb.Name+"'!"
 foreach($v in 'different_count','same_count','same_owner','retained'){ $g=$xl.Run($q+'IssueOutput',$v); Check ($g -like 'PASS*') "#1 output $v : $g" }
 foreach($v in 'model','field','shape','q','factor','incomplete','epoch'){ $g=$xl.Run($q+'IssueOutputGuard',$v); Check ($g -like 'ERROR 5:*') "#1 guard $v : $g" }
 foreach($line in (Get-Content "$P\tokens.txt" -Encoding UTF8)){ if(-not $line){continue}; $p=$line.Split("`t"); $txt=$p[1].Replace('\n',"`n").Replace('\r',"`r").Replace('\t',"`t"); $g=[bool]$xl.Run($q+'IssueEligible',[int]$p[0],$txt,$false,$false); Check ($g -eq [bool][int]$p[2]) "#2 token [$($p[0]) $($p[1])] expected=$([bool][int]$p[2]) actual=$g" }
 foreach($ca in @(@($true,$false),@($false,$true))){ $g=[bool]$xl.Run($q+'IssueEligible',5,'NEWTON_FACTORIZATION_FAILED pivot=123',$ca[0],$ca[1]); Check (-not $g) "#2 cancel=$($ca[0]) audit=$($ca[1]) -> $g" }
 $exp=@{ normal='engine=LEGACY;rescues=0;hsd_calls=0;status=OPTIMAL'; factor='engine=HSD_RESCUE;rescues=1;hsd_calls=1;status=OPTIMAL' }
 foreach($v in 'normal','factor','memory','subscript','cancel','unknown','audit','model','hsd_failure'){
  $g=$xl.Run($q+'IssueSolve',$v,"$Out\issueprobe_${Tag}_$v")
  switch($v){ 'normal'{$ok=$g.Contains($exp.normal)} 'factor'{$ok=$g.Contains($exp.factor)} 'memory'{$ok=$g.StartsWith('ERROR 7:') -and $g.Contains('hsd_calls=0')} 'subscript'{$ok=$g.StartsWith('ERROR 9:') -and $g.Contains('hsd_calls=0')} 'cancel'{$ok=$g.StartsWith('ERROR 18:') -and $g.Contains('hsd_calls=0')} 'hsd_failure'{$ok=$g.StartsWith('ERROR 7:HSD_TEST_MEMORY') -and $g.Contains('hsd_calls=1')} default{$ok=$g.StartsWith('ERROR 5:') -and $g.Contains('hsd_calls=0')} }
  Check $ok "#2 solver $v : $g" }
 $g=$xl.Run($q+'SmallRoots',"$Out\issueprobe_${Tag}_roots"); Check ($g -like 'PASS*') "roots : $g"
 Log "SUMMARY fail=$fail"
} catch { Log "ERR $_"; Log "SUMMARY fail=ERR" } finally { EndXL $xl $wb }
