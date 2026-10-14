param([Parameter(Mandatory)][string]$Kit,[Parameter(Mandatory)][string]$FixBook,[Parameter(Mandatory)][string]$Out,[string]$Only='')
# Sequential: exactly one owned Excel at a time. Skips runs whose result.json already exists.
$ErrorActionPreference='Stop'
$base=Join-Path $Kit 'RPFEM_20261014_G1R18.xlsm'
$runner=Join-Path $Kit 'run_issue12_case.ps1'
$common='ADAPT=1,QUADRATURE=8,RAW_LOWER_AUDIT=1'
$cases=[ordered]@{}
foreach($v in 'base','fix'){
  foreach($c in @(@('acc',1,0),@('fixed',0,1),@('off',0,0))){
    foreach($cache in 1,0){
      $cases["${v}_std_$($c[0])_cache$cache"]=@{book=$v;model='';set="TARGET=512,CYCLES=1,$common,LOWER_MODEL_CACHE=$cache,ACCURACY_DLL_RECOVERY=$($c[1]),LOWER_FIXED_RECOVERY=$($c[2])"}
    }
  }
  $cases["${v}_D07A_default"]=@{book=$v;model=(Join-Path $Kit 'D07A.json');set=''}
  $cases["${v}_std_default"]=@{book=$v;model='';set=''}
}
foreach($name in $cases.Keys){
  if($Only -and $name -notmatch $Only){continue}
  $d=Join-Path $Out $name
  if(Test-Path (Join-Path $d 'result.json')){ Write-Output "skip $name"; continue }
  while(Get-Process EXCEL -ErrorAction SilentlyContinue){ Write-Output 'waiting for other Excel'; Start-Sleep 30 }
  New-Item -ItemType Directory -Force $d | Out-Null
  Copy-Item (Join-Path $Kit '*.dll') $d
  $c=$cases[$name]; $src= if($c.book -eq 'base'){$base}else{$FixBook}
  $a=@('-NoProfile','-ExecutionPolicy','Bypass','-File',$runner,'-SourceBook',$src,'-Name',("RPFEM_"+($name -replace '_','-')),'-RunDir',$d,'-ProbePath',(Join-Path $Kit 'Issue12Probe.bas'))
  if($c.model){$a+=@('-ModelJson',$c.model)}
  if($c.set){$a+=@('-Settings',$c.set)}
  & powershell @a
}
Write-Output 'MATRIX DONE'
