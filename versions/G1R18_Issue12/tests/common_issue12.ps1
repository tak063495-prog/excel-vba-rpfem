# Shared helpers. Only Excel processes started here are ever closed.
Add-Type -Namespace Issue12 -Name Win -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(System.IntPtr h, out uint p);'
function SourceText($path){ (([IO.File]::ReadAllText($path,[Text.Encoding]::UTF8) -split '\r?\n' | Where-Object {$_ -notlike 'Attribute VB_*'}) -join "`r`n").Trim() }
function Start-OwnedExcel {
    $running=@(Get-Process EXCEL -ErrorAction SilentlyContinue)
    if($running.Count -gt 0){ throw "Another Excel is running (PIDs $($running.Id -join ',')); refusing to start a second Excel" }
    $xl=New-Object -ComObject Excel.Application
    [uint32]$p=0; [void][Issue12.Win]::GetWindowThreadProcessId([IntPtr]$xl.Hwnd,[ref]$p)
    $xl.Visible=$false; $xl.DisplayAlerts=$false
    return @{xl=$xl;pid=[int]$p}
}
function Stop-OwnedExcel($h,$book){
    try{ if($book){$book.Close($false)} }catch{}
    try{ $h.xl.Quit() }catch{}
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($h.xl); [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    Start-Sleep 3
    $proc=Get-Process -Id $h.pid -ErrorAction SilentlyContinue
    if($proc){ Stop-Process -Id $h.pid -Force }   # only the PID this script started
}
function Compile-Book($xl,$book,$module){
    $book.Activate(); $book.VBProject.VBComponents.Item($module).CodeModule.CodePane.Show()
    $ctl=$xl.VBE.CommandBars.FindControl(1,578)
    if($ctl.Enabled){$ctl.Execute()}
    if($ctl.Enabled){throw 'VBA compile failed'}
    $xl.VBE.MainWindow.Visible=$false
}
