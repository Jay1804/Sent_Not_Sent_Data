$ErrorActionPreference = 'Stop'
$src = $args[0]; $out = $args[1]
$x = New-Object -ComObject Excel.Application; $x.Visible = $false; $x.DisplayAlerts = $false
$lines = @()
try {
    $null = $x.Workbooks.Add(); $x.Calculation = -4135
    $w = $x.Workbooks.Open($src, 0, $true)
    $d = $w.Worksheets.Item("Sent Cases (CTS) - Bridge")
    $n = $d.UsedRange.Rows.Count
    $ids = $d.Range($d.Cells(2, 5), $d.Cells($n, 5)).Value2
    $sent = $d.Range($d.Cells(2, 33), $d.Cells($n, 33)).Value2
    $ai = $d.Range($d.Cells(2, 74), $d.Cells($n, 74)).Value2
    $cnt = @{}
    for ($i = 1; $i -le ($n - 1); $i += 97) {
        $v = if ($ai[$i, 1] -eq $null) { 'blank' } else { [string]$ai[$i, 1] }
        $cnt[$v] = 1 + [int]$cnt[$v]
        if ($cnt[$v] -le 25) { $lines += "$v,$($ids[$i,1]),$($sent[$i,1])" }
    }
    $w.Close($false)
} catch { $lines += "ERROR: $_" } finally { $x.Quit(); $lines | Set-Content $out -Encoding UTF8 }
