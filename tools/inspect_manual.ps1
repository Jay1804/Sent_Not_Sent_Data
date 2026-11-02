$ErrorActionPreference = 'Stop'
$src = $args[0]; $out = $args[1]
$x = New-Object -ComObject Excel.Application; $x.Visible = $false; $x.DisplayAlerts = $false
$lines = @()
try {
    $null = $x.Workbooks.Add(); $x.Calculation = -4135
    $w = $x.Workbooks.Open($src, 0, $true)
    $ws = $w.Worksheets.Item("Manual sentcount")
    $lines += "Used: " + $ws.UsedRange.Address()
    $b1 = $ws.Range("B1")
    $lines += "B1 value=" + $b1.Value2 + " formula=" + $b1.Formula
    try { $v = $b1.Validation; $lines += "B1 validation type=$($v.Type) f1=$($v.Formula1)" } catch { $lines += "B1 no validation: $_" }
    foreach ($r in 1..8) {
        foreach ($c in 1..6) {
            $cell = $ws.Cells($r, $c)
            if ($cell.Formula -ne "") { $lines += "R${r}C${c} [" + $cell.Address($false, $false) + "] formula=" + $cell.Formula + " | value=" + $cell.Text + " | fmt=" + $cell.NumberFormat }
        }
    }
    $lines += "last cols:"
    foreach ($r in 1..6) {
        foreach ($c in 32..36) {
            $cell = $ws.Cells($r, $c)
            if ($cell.Formula -ne "") { $lines += "R${r}C${c} formula=" + $cell.Formula + " | value=" + $cell.Text }
        }
    }
    $lines += "Names:"
    foreach ($n in $w.Names) { $lines += "  " + $n.Name + " -> " + $n.RefersTo }
    $w.Close($false)
} catch { $lines += "ERROR: $_" } finally { $x.Quit(); $lines | Set-Content $out -Encoding UTF8 }
