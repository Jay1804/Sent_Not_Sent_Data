$ErrorActionPreference = 'Stop'
$src = $args[0]; $out = $args[1]
$x = New-Object -ComObject Excel.Application; $x.Visible = $false; $x.DisplayAlerts = $false
$lines = @()
try {
    $null = $x.Workbooks.Add(); $x.Calculation = -4135
    $w = $x.Workbooks.Open($src, 0, $true)
    $d = $w.Worksheets.Item("Sent Cases (CTS) - Bridge")
    $col = 0
    for ($c = 1; $c -le 90; $c++) { if ([string]$d.Cells(1, $c).Value2 -eq 'Report AI Status') { $col = $c } }
    $lines += "col=$col"
    $c1 = $d.Cells(2, $col); $lines += "row2 formula=[" + $c1.Formula + "] fmt=" + $c1.NumberFormat
    $n = $d.UsedRange.Rows.Count
    $vals = $d.Range($d.Cells(2, $col), $d.Cells($n, $col)).Value2
    $h = @{}
    foreach ($v in $vals) { $k = if ($v -eq $null) { '<blank>' } else { [string]$v + ' (' + $v.GetType().Name + ')' }; $h[$k] = 1 + [int]$h[$k] }
    foreach ($k in $h.Keys) { $lines += "value $k = $($h[$k])" }
    # a few non-blank formulas
    $cnt = 0
    for ($r = 2; $r -le $n -and $cnt -lt 3; $r++) { $cell = $d.Cells($r, $col); if ($cell.Formula -ne '') { $lines += "R$r formula=[" + $cell.Formula + "] text=" + $cell.Text + " | status=" + $d.Cells($r, 39).Text; $cnt++ } }
    $w.Close($false)
} catch { $lines += "ERROR: $_" } finally { $x.Quit(); $lines | Set-Content $out -Encoding UTF8 }
