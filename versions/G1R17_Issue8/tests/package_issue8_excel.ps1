param([string]$Root = (Split-Path $PSScriptRoot))
$ErrorActionPreference='Stop'
Add-Type -Namespace Issue8Package -Name Win -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(System.IntPtr h, out uint p);'
$Root=[IO.Path]::GetFullPath($Root)
$Base=Join-Path (Split-Path $Root) 'G1R17'
$dest=Join-Path $Root 'delivery'
New-Item -ItemType Directory -Force $dest | Out-Null
$before=@(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
$xl=$null; $book=$null; $owned=$false; $results=@()
function SourceText($path){
    return (([IO.File]::ReadAllText($path,[Text.Encoding]::UTF8) -split '\r?\n' | Where-Object {$_ -notlike 'Attribute VB_*'}) -join "`r`n").Trim()
}
function Compile {
    $book.Activate(); $book.VBProject.VBComponents.Item('RPX_Seep').CodeModule.CodePane.Show()
    $ctl=$xl.VBE.CommandBars.FindControl(1,578)
    if($ctl.Enabled){$ctl.Execute()}
    if($ctl.Enabled){throw 'Compile failed'}
    $xl.VBE.MainWindow.Visible=$false
}
try {
    $xl=New-Object -ComObject Excel.Application
    [uint32]$excelPid=0
    [void][Issue8Package.Win]::GetWindowThreadProcessId([IntPtr]$xl.Hwnd,[ref]$excelPid)
    if($before -contains [int]$excelPid){throw 'Refusing existing Excel process'}
    $owned=$true; $xl.Visible=$false; $xl.DisplayAlerts=$false; $xl.EnableEvents=$false
    Copy-Item -LiteralPath (Join-Path $Base 'delivery/rpfem_native_x64.dll') -Destination $dest
    foreach($case in 'P06A','P06N'){
        $original=(Get-ChildItem (Join-Path $Base 'delivery') -Filter "*$case*.xlsm" | Select-Object -First 1).FullName
        $candidate=Join-Path $dest "RPFEM_G1R17_Issue8_${case}_Checked.xlsm"
        Copy-Item -LiteralPath $original -Destination $candidate
        $book=$xl.Workbooks.Open($candidate,0,$false)
        $baseline=@{}
        foreach($c in $book.VBProject.VBComponents){if($c.CodeModule.CountOfLines){$baseline[$c.Name]=$c.CodeModule.Lines(1,$c.CodeModule.CountOfLines)}else{$baseline[$c.Name]=''}}
        foreach($name in 'RPX_Seep','RPX_SeepUI'){
            $cm=$book.VBProject.VBComponents.Item($name).CodeModule
            $cm.DeleteLines(1,$cm.CountOfLines); $cm.AddFromString((SourceText (Join-Path $Root "src/vba/$name.bas")))
        }
        Compile
        $book.Save(); $book.Close($false); $book=$null
        $book=$xl.Workbooks.Open($candidate,0,$true)
        $count=$book.VBProject.VBComponents.Count
        if($count -ne $baseline.Count){throw 'Module count changed'}
        $casingOnly=@()
        foreach($c in $book.VBProject.VBComponents){
            $actual=''; if($c.CodeModule.CountOfLines){$actual=$c.CodeModule.Lines(1,$c.CodeModule.CountOfLines)}
            if($c.Name -in 'RPX_Seep','RPX_SeepUI'){
                if($actual.Trim() -ine (SourceText (Join-Path $Root "src/vba/$($c.Name).bas"))){throw "Source mismatch: $($c.Name)"}
            }elseif($actual -cne $baseline[$c.Name]){
                if($actual -ine $baseline[$c.Name]){throw "Unintended module change: $($c.Name)"}
                $casingOnly+=$c.Name
            }
        }
        # Probe saved production source in memory, then discard the test module.
        $c=$book.VBProject.VBComponents.Add(1); $c.Name='Issue8Probe'
        $c.CodeModule.AddFromString((SourceText (Join-Path $PSScriptRoot 'Issue8Probe.bas')))
        Compile
        $result=$xl.Run("'"+$book.Name+"'!Issue8Hydro")
        if(-not $result.StartsWith('True;1;')){throw 'Reopened hydro regression failed'}
        $results+=@{case=$case;compile='PASS';module_count=$count;source_match=$true;other_modules_equivalent=$true;casing_only_changes=$casingOnly;reopened_hydro=$result;sha256=(Get-FileHash -LiteralPath $candidate -Algorithm SHA256).Hash.ToLower();source_sha256=(Get-FileHash -LiteralPath $original -Algorithm SHA256).Hash.ToLower()}
        $book.Close($false); $book=$null
    }
    $results | ConvertTo-Json -Depth 6 | Set-Content -Encoding utf8 (Join-Path $Root 'results/issue8_package.json')
    Write-Output 'PACKAGE COMPILE REOPEN SOURCE MATCH PASS'
} finally {
    if($owned){if($book){$book.Close($false)}; if($xl){$xl.Quit();[void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)}}
}
