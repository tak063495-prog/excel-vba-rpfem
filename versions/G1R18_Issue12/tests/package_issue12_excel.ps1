param([Parameter(Mandatory)][string]$BaseBook,[Parameter(Mandatory)][string]$SrcDir,[Parameter(Mandatory)][string]$OutBook,[Parameter(Mandatory)][string]$ResultJson)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'common_issue12.ps1')
if(Test-Path -LiteralPath $OutBook){throw "Refusing to overwrite $OutBook"}
Copy-Item -LiteralPath $BaseBook -Destination $OutBook
$h=Start-OwnedExcel; $book=$null; $r=[ordered]@{excel_pid=$h.pid;base_sha256=(Get-FileHash $BaseBook).Hash.ToLower()}
try{
    $xl=$h.xl; $xl.EnableEvents=$false
    $book=$xl.Workbooks.Open($OutBook,0,$false)
    $baseline=@{}
    foreach($c in $book.VBProject.VBComponents){ $baseline[$c.Name]= if($c.CodeModule.CountOfLines){$c.CodeModule.Lines(1,$c.CodeModule.CountOfLines)}else{''} }
    $changed=@('RPX_LowerCache','RPX_Fs')
    foreach($n in $changed){ $cm=$book.VBProject.VBComponents.Item($n).CodeModule; $cm.DeleteLines(1,$cm.CountOfLines); $cm.AddFromString((SourceText (Join-Path $SrcDir "$n.bas"))) }
    Compile-Book $xl $book 'RPX_LowerCache'
    $book.Save(); $book.Close($false); $book=$null
    $book=$xl.Workbooks.Open($OutBook,0,$true)
    if($book.VBProject.VBComponents.Count -ne $baseline.Count){throw 'Module count changed'}
    $casing=@()
    foreach($c in $book.VBProject.VBComponents){
        $a= if($c.CodeModule.CountOfLines){$c.CodeModule.Lines(1,$c.CodeModule.CountOfLines)}else{''}
        if($c.Name -in $changed){ if($a.Trim() -ine (SourceText (Join-Path $SrcDir "$($c.Name).bas"))){throw "Source mismatch $($c.Name)"} }
        elseif($a -cne $baseline[$c.Name]){ if($a -ine $baseline[$c.Name]){throw "Unintended change $($c.Name)"}; $casing+=$c.Name }
    }
    Compile-Book $xl $book 'RPX_LowerCache'
    $r.module_count=$book.VBProject.VBComponents.Count; $r.compile='PASS'; $r.reopen='PASS'; $r.source_match=$true; $r.casing_only=$casing
    $book.Close($false); $book=$null
    $r.sha256=(Get-FileHash $OutBook).Hash.ToLower(); $r.status='PASS'
} catch { $r.status='FAIL'; $r.error="$($_.Exception.Message)"; throw }
finally { $r | ConvertTo-Json -Depth 5 | Set-Content -Encoding utf8 $ResultJson; Stop-OwnedExcel $h $book }
