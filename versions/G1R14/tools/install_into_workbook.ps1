# install_into_workbook.ps1
# Installs RPX_Seep / RPX_SeepUI / RPX_Water into an Excel workbook and inserts the two hook lines.
# Requirement: Excel Trust Center -> "Trust access to the VBA project object model" = ON.
# Usage: powershell -ExecutionPolicy Bypass -File install_into_workbook.ps1 -Workbook "C:\path\tool.xlsm"
param(
    [Parameter(Mandatory = $true)][string]$Workbook,
    [switch]$SkipSetupSheet
)
$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$wbPath = (Resolve-Path $Workbook).Path
$modules = @('RPX_Seep', 'RPX_SeepUI', 'RPX_Water')
foreach ($m in $modules) {
    if (-not (Test-Path (Join-Path $here "$m`_cp932.bas"))) { throw "Missing $m`_cp932.bas next to this script." }
}

# 1. backup
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$backup = [IO.Path]::Combine([IO.Path]::GetDirectoryName($wbPath), [IO.Path]::GetFileNameWithoutExtension($wbPath) + "_backup_$stamp" + [IO.Path]::GetExtension($wbPath))
Copy-Item $wbPath $backup
Write-Host "Backup: $backup"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.EnableEvents = $false
try {
    # 2. remove old modules, save, close (avoids RPX_Seep1 style renaming on re-import)
    $wb = $excel.Workbooks.Open($wbPath)
    try { $null = $wb.VBProject.VBComponents.Count } catch { throw "VBA project access denied. Enable 'Trust access to the VBA project object model'." }
    foreach ($m in $modules) {
        foreach ($c in @($wb.VBProject.VBComponents)) {
            if ($c.Name -eq $m -or $c.Name -like "$m[0-9]*") {
                Write-Host "Remove $($c.Name)"
                $wb.VBProject.VBComponents.Remove($c)
            }
        }
    }
    $wb.Save(); $wb.Close($true)

    # 3. reopen and import
    $wb = $excel.Workbooks.Open($wbPath)
    foreach ($m in $modules) {
        $comp = $wb.VBProject.VBComponents.Import((Join-Path $here "$m`_cp932.bas"))
        Write-Host "Imported $($comp.Name)"
        if ($comp.Name -ne $m) { throw "Imported as $($comp.Name) instead of $m. Save/reopen and retry." }
    }

    # 4. hook lines (idempotent)
    function Insert-Hook($wb, $moduleName, $procName, $anchor, $newLine) {
        $cm = $wb.VBProject.VBComponents.Item($moduleName).CodeModule
        $first = $cm.ProcStartLine($procName, 0)
        $count = $cm.ProcCountLines($procName, 0)
        $hit = 0
        for ($i = $first; $i -lt $first + $count; $i++) {
            $t = $cm.Lines($i, 1).Trim()
            if ($t -eq $newLine.Trim()) { Write-Host "$moduleName.$procName : already present (line $i)"; return }
            if ($t -eq $anchor -and $hit -eq 0) { $hit = $i }
        }
        if ($hit -eq 0) { throw "$moduleName.$procName : anchor not found [$anchor]" }
        $cm.InsertLines($hit + 1, $newLine)
        Write-Host "$moduleName.$procName : inserted at line $($hit + 1)"
    }
    # G1R13+: insert after the load-step repair line (RPX_LoadPhase="raw" moved before RPX_Topology).
    $meshCm = $wb.VBProject.VBComponents.Item('RPX_Mesh').CodeModule
    $meshText = $meshCm.Lines(1, $meshCm.CountOfLines)
    if ($meshText.Contains('If RPX_RepairLoadSteps() Then Call RPX_AssignRegions')) {
        Insert-Hook $wb 'RPX_Mesh' 'RPX_AssignRegions' 'If RPX_RepairLoadSteps() Then Call RPX_AssignRegions' '    Call RPX_WaterApply'
    } else {
        Insert-Hook $wb 'RPX_Mesh' 'RPX_AssignRegions' 'RPX_LoadPhase = "raw"' '    Call RPX_WaterApply'
    }
    Insert-Hook $wb 'RPX_Control' 'RPX_DrawCore' 'If RPX_ResultCurrent Then Call RPX_DrawVelocity' '    Call RPX_SeepOverlay'
    $wb.Save()

    # 5. create/format the seepage input sheet
    if (-not $SkipSetupSheet) {
        $excel.EnableEvents = $true
        $excel.Run("'" + $wb.Name + "'!RPX_SeepSetupSheet")
        $wb.Save()
        Write-Host "RPX_SeepSetupSheet done"
    }
    $wb.Close($true)
    Write-Host "Done. Open the workbook, run Debug > Compile VBAProject, then ?RPX_SeepOverlaySelfTest in the Immediate window."
}
finally {
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
