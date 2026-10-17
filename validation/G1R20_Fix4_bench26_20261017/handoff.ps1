# One-shot hand-off: wait until the old batch spawns its next worker (which idles because HANDOFF exists),
# then terminate only the bench runner python processes, and restart the main task with the revised timeouts.
$root='C:\Users\link_\rpfem_bench\validation\G1R20_Fix4_bench26'
$log="$root\results\handoff_log.txt"
function L($m){ Add-Content -Path $log -Value ("{0} {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$m) -Encoding UTF8 }
L 'watcher start'
$deadline=(Get-Date).AddHours(10)
while(-not (Get-ChildItem "$root\results\handoff_worker_*.json" -ErrorAction SilentlyContinue)){
  if((Get-Date) -gt $deadline){ L 'gave up waiting'; exit 1 }
  if((Get-ScheduledTask RPFEM_G1R20_Fix4_bench26).State -ne 'Running'){ L 'main task not running while waiting'; break }
  Start-Sleep 2
}
$w=Get-ChildItem "$root\results\handoff_worker_*.json" | Select -First 1
L ("handoff worker seen: " + $w.Name)
$procs=Get-CimInstance Win32_Process -Filter "Name='python.exe'" | ? { $_.CommandLine -match 'runner\.py (batch|case)' -and ($_.CommandLine -match 'G1R20_Fix4_bench26' -or $_.CommandLine -match 'scripts\\runner\.py batch') }
foreach($p in $procs){ L ("kill python pid {0}: {1}" -f $p.ProcessId,$p.CommandLine); Stop-Process -Id $p.ProcessId -Force -ErrorAction SilentlyContinue }
$t=(Get-Date).AddMinutes(3)
while((Get-ScheduledTask RPFEM_G1R20_Fix4_bench26).State -eq 'Running' -and (Get-Date) -lt $t){ Start-Sleep 2 }
L ("main task state: " + (Get-ScheduledTask RPFEM_G1R20_Fix4_bench26).State)
Remove-Item "$root\HANDOFF" -Force
Get-ChildItem "$root\results\handoff_worker_*.json" | % { Move-Item $_.FullName ($_.FullName + '.done') -Force }
Start-ScheduledTask -TaskName RPFEM_G1R20_Fix4_bench26
L 'main task restarted with revised timeouts'
