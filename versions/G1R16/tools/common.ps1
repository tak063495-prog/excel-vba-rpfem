# Parallel-safe: never touches Excel processes this script did not create.
Add-Type -Namespace W -Name U -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(System.IntPtr h, out uint p);'
$script:Protected=@(31656)
function NewLog($p){ $script:LogPath=$p }
function Log($m){ Add-Content $script:LogPath ((Get-Date -f HH:mm:ss)+' '+$m) -Encoding UTF8 }
function Sheet($wb,$n){ foreach($sh in $wb.Sheets){ if($sh.Name -eq $n){ return $sh } } }
function SetV($sh,$a,$v){ $r=$sh.Range($a); [void]$r.GetType().InvokeMember('Value2',[Reflection.BindingFlags]::SetProperty,$null,$r,@($v)) }
function SetC($c,$v){ [void]$c.GetType().InvokeMember('Value2',[Reflection.BindingFlags]::SetProperty,$null,$c,@($v)) }
function StartXL(){
  $before=@(Get-Process EXCEL -ea 0 | % Id)
  $x=New-Object -ComObject Excel.Application; $x.Visible=$false; $x.DisplayAlerts=$false; $x.AutomationSecurity=1
  [uint32]$p=0; [void][W.U]::GetWindowThreadProcessId([IntPtr]$x.Hwnd,[ref]$p); $script:MyPid=[int]$p
  if($before -contains $script:MyPid -or $script:Protected -contains $script:MyPid){ Log "ABORT: COM attached to existing Excel pid $($script:MyPid)"; $script:MyPid=0; throw 'ATTACHED_TO_FOREIGN_EXCEL' }
  Log "own Excel pid=$($script:MyPid)"; return $x }
function EndXL($x,$wb){
  try{ if($wb){ $wb.Close($false) } }catch{}; try{ $x.Quit() }catch{}
  try{ [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }catch{}; [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  if($script:MyPid -gt 0 -and $script:Protected -notcontains $script:MyPid){
    for($i=0;$i -lt 30 -and (Get-Process -Id $script:MyPid -ea 0);$i++){ Start-Sleep 1 }
    if(Get-Process -Id $script:MyPid -ea 0){ Stop-Process -Id $script:MyPid -Force; Log "own Excel $($script:MyPid) force-closed after quit" } } }
function Compile($xl,$wb){
  $wb.Activate(); $c=$wb.VBProject.VBComponents.Item('RPX_Fs'); $c.CodeModule.CodePane.Show()
  $ctl=$xl.VBE.CommandBars.FindControl(1,578)
  if($ctl.Enabled){ $ctl.Execute(); Start-Sleep -Milliseconds 500 }
  $ok= -not $ctl.Enabled; $xl.VBE.MainWindow.Visible=$false; return $ok }
