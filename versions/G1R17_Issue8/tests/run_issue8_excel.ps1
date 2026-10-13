param([string]$Root = (Split-Path $PSScriptRoot), [string]$RebuiltDll)
$ErrorActionPreference='Stop'
Add-Type -Namespace Issue8 -Name Win -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(System.IntPtr h, out uint p);'
$Root=[IO.Path]::GetFullPath($Root)
$Base=Join-Path (Split-Path $Root) 'G1R17'
$out=Join-Path $Root 'results/issue8_excel'
New-Item -ItemType Directory -Force $out | Out-Null
$results=[ordered]@{}
$before=@(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
$xl=$null; $book=$null; $owned=$false
function InvokeMacro($name,$arguments) {
    $args=@("'"+$book.Name+"'!"+$name)+$arguments
    return $xl.GetType().InvokeMember('Run',[Reflection.BindingFlags]::InvokeMethod,$null,$xl,$args)
}
function ReplaceModule($name,$path) {
    $cm=$book.VBProject.VBComponents.Item($name).CodeModule
    $s=[IO.File]::ReadAllText($path,[Text.Encoding]::UTF8)
    $s=($s -split '\r?\n' | Where-Object {$_ -notlike 'Attribute VB_*'}) -join "`r`n"
    $cm.DeleteLines(1,$cm.CountOfLines); $cm.AddFromString($s)
}
try {
    $xl=New-Object -ComObject Excel.Application
    [uint32]$excelPid=0
    [void][Issue8.Win]::GetWindowThreadProcessId([IntPtr]$xl.Hwnd,[ref]$excelPid)
    if($before -contains [int]$excelPid){throw 'Refusing to use an existing Excel process'}
    $owned=$true
    $xl.Visible=$false; $xl.DisplayAlerts=$false; $xl.EnableEvents=$false
    # Existing macro and VBA access policies are preserved.
    $results['environment']=[ordered]@{version=$xl.Version;pid=$excelPid;prior_pids=$before;automation_security=$xl.AutomationSecurity}
    $source=(Get-ChildItem (Join-Path $Base 'delivery') -Filter '*P06A*.xlsm' | Select-Object -First 1).FullName
    foreach($variant in 'baseline','patched') {
        $dir=Join-Path $out $variant
        New-Item -ItemType Directory -Force $dir | Out-Null
        $candidate=Join-Path $out "$variant-test.xlsm"
        Copy-Item -LiteralPath $source -Destination $candidate
        $dll=Join-Path $Base 'delivery/rpfem_native_x64.dll'
        if($RebuiltDll){$dll=[IO.Path]::GetFullPath($RebuiltDll)}
        if($variant -eq 'baseline'){Copy-Item -LiteralPath $dll -Destination (Join-Path $out 'rpfem_native_x64.dll')}
        $book=$xl.Workbooks.Open($candidate,0,$false)
        if($variant -eq 'patched'){
            foreach($n in 'RPX_Seep','RPX_SeepUI'){ReplaceModule $n (Join-Path $Root "src/vba/$n.bas")}
        }
        $c=$book.VBProject.VBComponents.Add(1); $c.Name='Issue8Probe'
        $s=[IO.File]::ReadAllText((Join-Path $PSScriptRoot 'Issue8Probe.bas'),[Text.Encoding]::UTF8)
        $c.CodeModule.AddFromString((($s -split '\r?\n' | Where-Object {$_ -notlike 'Attribute VB_*'}) -join "`r`n"))
        $book.Activate(); $book.VBProject.VBComponents.Item('RPX_Seep').CodeModule.CodePane.Show()
        $ctl=$xl.VBE.CommandBars.FindControl(1,578)
        if($ctl.Enabled){$ctl.Execute()}
        if($ctl.Enabled){throw "VBA compilation failed: $variant"}
        $xl.VBE.MainWindow.Visible=$false
        $results[$variant]=[ordered]@{compile='PASS';runs=[ordered]@{}}
        foreach($mode in 'vba','auto'){
            InvokeMacro 'RPX_NativeSetMode' @($mode) | Out-Null
            $r=[ordered]@{native_info=(InvokeMacro 'RPX_NativeInfo' @());hydro=(InvokeMacro 'Issue8Hydro' @());relax_normal=(InvokeMacro 'RPX_SeepTestRelax' @(0.5));relax_tiny=(InvokeMacro 'RPX_SeepTestRelax' @(1e-8))}
            if($mode -eq 'auto' -and -not $r.native_info.StartsWith('DLL_READY;')){throw "DLL unavailable: $($r.native_info)"}
            foreach($dims in @(@(12,6),@(20,10),@(40,20))){
                $r["dam_$($dims[0])x$($dims[1])"]=InvokeMacro 'Issue8Dam' @([int]$dims[0],[int]$dims[1])
            }
            if($variant -eq 'patched'){
                $passed=0
                foreach($datum in -10.,10.,1e6){foreach($kval in 1e-12,1.,1e12){foreach($turn in 0,1,2){foreach($multi in $false,$true){
                    $v=InvokeMacro 'Issue8Hydro' @([double]$datum,[double]$kval,[int]$turn,[bool]$multi)
                    if(-not $v.StartsWith('True;')){throw "Hydro variant failed: $datum $kval $turn $multi : $v"}
                    $passed++
                }}}}
                $r['hydro_variants_pass']=$passed
                $r['coupling_rejected']=InvokeMacro 'Issue8CouplingReject' @()
                if(-not $r['coupling_rejected']){throw 'Unconverged coupling was not rejected'}
            }
            $results[$variant]['runs'][$mode]=$r
            Write-Output "$variant / $mode completed: hydro=$($r.hydro) native=$($r.native_info)"
        }
        $book.Save()
        $book.Close($false); $book=$null
        # Reopen, compare changed module source and execute saved patched solver.
        $book=$xl.Workbooks.Open($candidate,0,$true)
        if($variant -eq 'patched'){
            foreach($n in 'RPX_Seep','RPX_SeepUI'){
                $cm=$book.VBProject.VBComponents.Item($n).CodeModule
                $actual=$cm.Lines(1,$cm.CountOfLines).Replace("`r`n","`n").Trim()
                [IO.File]::WriteAllText((Join-Path $out "saved_$n.bas"),$actual,[Text.Encoding]::UTF8)
                $expected=([IO.File]::ReadAllText((Join-Path $Root "src/vba/$n.bas"),[Text.Encoding]::UTF8) -split '\r?\n' | Where-Object {$_ -notlike 'Attribute VB_*'}) -join "`n"
                # VBE changes identifier casing and literal presentation.
                $results[$variant]["saved_$n"]=$actual -ieq $expected.Trim()
                if(-not $results[$variant]["saved_$n"]){throw "Saved source differs: $n"}
            }
            $results[$variant]['reopened_hydro']=InvokeMacro 'Issue8Hydro' @()
        }
        $book.Close($false); $book=$null
    }
    foreach($mode in 'vba','auto'){
        if(-not $results.patched.runs[$mode].hydro.StartsWith('True;1;')){throw 'Hydro must converge in one iteration'}
        if(-not $results.patched.runs[$mode].relax_tiny.Contains('converged=False')){throw '#6 regression'}
        foreach($shape in 'dam_12x6','dam_20x10','dam_40x20'){
            # Ignore residual metric (new meaning), compare report, iterations and all heads.
            $a=$results.baseline.runs[$mode][$shape].Split(';')
            $b=$results.patched.runs[$mode][$shape].Split(';')
            $a[4]=''; $b[4]=''
            if(($a -join ';') -ne ($b -join ';')){throw "Dam changed: $mode $shape"}
        }
    }
    $results['status']='PASS'
} catch {
    $results['status']='FAIL'; $results['error']="$($_.Exception.Message)"
    throw
} finally {
    $results | ConvertTo-Json -Depth 12 | Set-Content -Encoding utf8 (Join-Path $out 'results.json')
    if($owned){if($book){$book.Close($false)}; if($xl){$xl.Quit();[void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)}}
}
