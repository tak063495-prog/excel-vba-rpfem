$Global:RpxMarker=(Join-Path $env:TEMP 'RPFEM_RUNNING_ANALYSIS.marker')
# Never kill an Excel that is mid-analysis. Marker = "<excel pid> <tag> <time>".
function Stop-OrphanExcel {
  if(Test-Path $Global:RpxMarker){
    $p=0; [void][int]::TryParse(((Get-Content $Global:RpxMarker -TotalCount 1) -split ' ')[0],[ref]$p)
    for($i=0;$i -lt 24 -and $p -and (Get-Process -Id $p -ErrorAction SilentlyContinue);$i++){ Start-Sleep 5 }
    if($p -and (Get-Process -Id $p -ErrorAction SilentlyContinue)){ throw "GUARD: analysis Excel pid $p is running ($(Get-Content $Global:RpxMarker)). Refusing to kill." }
    Remove-Item $Global:RpxMarker -Force
  }
  if(Get-Process EXCEL -ErrorAction SilentlyContinue){ throw 'Excel is running. Close it yourself and retry; this script never kills Excel.' }
}
function Set-RunMarker($tag){ $x=Get-Process EXCEL -ErrorAction SilentlyContinue | Sort-Object StartTime -Descending | Select-Object -First 1; if($x){ Set-Content $Global:RpxMarker "$($x.Id) $tag $(Get-Date -f s)" } }
function Clear-RunMarker { Remove-Item $Global:RpxMarker -Force -ErrorAction SilentlyContinue }
