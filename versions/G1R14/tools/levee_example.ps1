# levee_example.ps1  - 堤防（河川側外水位）例題：定常浸透＋連成FS
# Usage: powershell -ExecutionPolicy Bypass -File levee_example.ps1 -Workbook "C:\path\tool.xlsm"
param([Parameter(Mandatory = $true)][string]$Workbook)
$ErrorActionPreference = 'Stop'
$wbPath = (Resolve-Path $Workbook).Path
$dir = Split-Path -Parent $wbPath
$logPath = Join-Path $dir 'levee_run.log'
$pngPath = Join-Path $dir 'diagram_levee.png'
function Log([string]$msg) {
    $line = (Get-Date -Format 'HH:mm:ss') + ' ' + $msg
    Add-Content -Path $logPath -Value $line -Encoding UTF8
    Write-Host $line
}
function Arr2([object[]]$rows, [int]$cols) {
    $n = $rows.Count
    $a = New-Object 'object[,]' $n, $cols
    for ($i = 0; $i -lt $n; $i++) {
        for ($j = 0; $j -lt $cols; $j++) {
            $v = $rows[$i][$j]
            if ($v -is [int]) { $v = [double]$v }
            $a[$i, $j] = $v
        }
    }
    return ,$a
}
function SetBlock($ws, [string]$topLeft, [object[]]$rows, [int]$cols) {
    $arr = Arr2 $rows $cols
    $ws.Range($topLeft).Resize($rows.Count, $cols).Value2 = $arr
}
function Cell($ws, [string]$addr) { return $ws.Range($addr).Value2 }

Set-Content -Path $logPath -Value ((Get-Date -Format 'HH:mm:ss') + ' start ' + $wbPath) -Encoding UTF8
$excel = New-Object -ComObject Excel.Application
$excel.AutomationSecurity = 1
$excel.DisplayAlerts = $false
$excel.Visible = $true
$wb = $null
try {
    $wb = $excel.Workbooks.Open($wbPath)
    $book = "'" + $wb.Name + "'!"

    # ---- 要素定義：旧データ全消去 → 堤防形状
    $def = $wb.Worksheets.Item('要素定義')
    $def.Range('A7:C1006').ClearContents() | Out-Null
    $def.Range('E7:H1006').ClearContents() | Out-Null
    $def.Range('J7:T1006').ClearContents() | Out-Null
    $pts = @(
        @(1, 0, 0), @(2, 40, 0), @(3, 40, 2), @(4, 28, 2),
        @(5, 18, 7), @(6, 14, 7), @(7, 4, 2), @(8, 0, 2))
    SetBlock $def 'A7' $pts 3
    # 領域：ID, 材料ID, 剛体ID(0=なし), 境界点（定義順・反時計回り）
    SetBlock $def 'E7' @(,@(1, 1, 0, '1,2,3,4,5,6,7,8')) 4
    # 面：ID, 領域, 開始点, 終了点, 循環, 拘束, 剛体ID, 固定Fx, 固定Fy, 比例Fx, 比例Fy
    $faces = @(
        @(1, 1, 1, 2, 0, '固定XY', 0, 0, 0, 0, 0),
        @(2, 1, 2, 3, 0, '固定X', 0, 0, 0, 0, 0),
        @(3, 1, 3, 4, 0, '自由', 0, 0, 0, 0, 0),
        @(4, 1, 4, 5, 0, '自由', 0, 0, 0, 0, 0),
        @(5, 1, 5, 6, 0, '自由', 0, 0, 0, 0, 0),
        @(6, 1, 6, 7, 0, '自由', 0, 0, 0, 0, 0),
        @(7, 1, 7, 8, 0, '自由', 0, 0, 0, 0, 0),
        @(8, 1, 8, 1, 1, '固定X', 0, 0, 0, 0, 0))
    SetBlock $def 'J7' $faces 11
    Log 'geometry: 8 points, 1 region, 8 faces (bottom fixed XY, sides fixed X, surfaces free)'

    # ---- 材料データ：c, phi, psi は維持し、γ=19（飽和単位体積重量として扱う）
    $mat = $wb.Worksheets.Item('材料データ')
    $mat.Range('D5').Value2 = [double]19
    Log ('material 1: c=' + (Cell $mat 'B5') + ' phi=' + (Cell $mat 'C5') + ' gamma=19 (saturated) psi=' + (Cell $mat 'E5'))

    # ---- 浸透入力
    $excel.Run($book + 'RPX_SeepSetupSheet') | Out-Null
    $sp = $wb.Worksheets.Item('浸透入力')
    $sp.Range('A7:C1006').ClearContents() | Out-Null
    $sp.Range('E7:K1006').ClearContents() | Out-Null
    SetBlock $sp 'A7' @(,@(1, 1e-5, 1e-5)) 3
    # 境界：ID, 領域, 開始点, 終了点, 循環, 種別, 値
    $bcs = @(
        @(1, 1, 6, 8, 0, '水位', 6.0),
        @(2, 1, 8, 1, 1, '水位', 6.0),
        @(3, 1, 3, 5, 0, '浸出面', 0.0),
        @(4, 1, 2, 3, 0, '水位', 2.0))
    SetBlock $sp 'E7' $bcs 7
    $sp.Range('N19').Value2 = [double]1
    Log 'seepage input: k=1e-5 m/s; river side WL 6.0 (pt6->pt8->pt1); land slope/toe seepage face (pt3->pt5); right edge WL 2.0; coupling N19=1'

    # ---- 浸透解析
    $excel.Run($book + 'RPX_SeepRun') | Out-Null
    $res = $wb.Worksheets.Item('浸透結果')
    Log ('seepage: iterations=' + (Cell $res 'K1') + ' converged=' + (Cell $res 'K2') + ' lastDelta=' + (Cell $res 'K3') +
         ' inflow=' + (Cell $res 'K4') + ' outflow=' + (Cell $res 'K5') + ' elements=' + (Cell $res 'K6') + ' FSsegments=' + (Cell $res 'K8'))
    # 非水面（天端・裏法・裏側地表）で地表より水頭が高い節点の検査
    $nNodes = [int](Cell $res 'K7')
    $data = $res.Range('B2').Resize($nNodes, 4).Value2
    $bad = 0; $maxp = 0.0
    for ($i = 1; $i -le $nNodes; $i++) {
        $x = [double]$data[$i, 1]; $y = [double]$data[$i, 2]; $p = [double]$data[$i, 4]
        $yg = $null
        if ($x -ge 14 -and $x -le 18) { $yg = 7.0 }
        elseif ($x -gt 18 -and $x -le 28) { $yg = 7.0 - ($x - 18.0) / 2.0 }
        elseif ($x -gt 28 -and $x -le 40) { $yg = 2.0 }
        if ($yg -ne $null -and [Math]::Abs($y - $yg) -lt 1e-6 -and $p -gt 1e-4) {
            $bad++; if ($p -gt $maxp) { $maxp = $p }
        }
    }
    Log ('check: land/crest surface nodes with pressure head > 0: ' + $bad + ' (max ' + $maxp + ' m)')

    # ---- 極限解析（連成）
    $t0 = Get-Date
    try {
        $excel.Run($book + 'RPX_Run') | Out-Null
        Log ('RPX_Run returned after ' + [int]((Get-Date) - $t0).TotalSeconds + ' s')
    } catch {
        $err = $_.Exception.Message
        Log ('RPX_Run exception after ' + [int]((Get-Date) - $t0).TotalSeconds + ' s: ' + $err)
    }
    $ar = $wb.Worksheets.Item('解析結果')
    Log ('result A1: ' + (Cell $ar 'A1'))
    Log ('FS lower=' + (Cell $ar 'B4') + ' upper=' + (Cell $ar 'C4') + ' elements=' + (Cell $ar 'B10') + '/' + (Cell $ar 'C10'))

    # ---- 図
    try { $excel.Run($book + 'RPX_ShowDiagram') | Out-Null } catch { Log ('RPX_ShowDiagram exception: ' + $_.Exception.Message) }
    $fig = $wb.Worksheets.Item('図')
    Log ('diagram A4: ' + (Cell $fig 'A4'))
    Log ('diagram A5: ' + (Cell $fig 'A5'))
    Log ('diagram A6: ' + (Cell $fig 'A6'))
    try {
        $fig.Activate() | Out-Null
        $rng = $fig.Range('A1:Z40')
        $rng.CopyPicture(1, -4147) | Out-Null
        $co = $fig.ChartObjects().Add(0, 0, $rng.Width, $rng.Height)
        $co.Activate() | Out-Null
        $co.Chart.Paste() | Out-Null
        $co.Chart.Export($pngPath) | Out-Null
        $co.Delete() | Out-Null
        Log ('diagram exported: ' + $pngPath)
    } catch {
        Log ('diagram export failed: ' + $_.Exception.Message)
    }
    $wb.Save()
    Log 'saved'
}
catch {
    Log ('ERROR: ' + $_.Exception.Message)
}
finally {
    Log 'end'
}
