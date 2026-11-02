# usage: make_workbook.ps1 <dump.csv> <meta.json> <out> <fileformat>
# Excel imports the CSV natively (fast), then we add the other tabs, the
# Manual sentcount formulas (with the cluster drop-down) and real PivotTables.
$ErrorActionPreference = 'Stop'
$csv = $args[0]; $metaJson = $args[1]; $out = $args[2]; $fmt = [int]$args[3]
$sw = [Diagnostics.Stopwatch]::StartNew()
function T($m) { Write-Host ("[{0:N0}s] {1}" -f $sw.Elapsed.TotalSeconds, $m) }
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false; $x.ScreenUpdating = $false
try {
    $wb = $x.Workbooks.Add(); $x.Calculation = -4135
    $meta = Get-Content $metaJson -Raw | ConvertFrom-Json

    # QueryTable import: honours per-column types (text for IDs/comments, General for the rest)
    $d = $wb.Worksheets.Item(1)
    $types = New-Object 'int[]' ([int]$meta.ncols)
    for ($i = 0; $i -lt $types.Length; $i++) { $types[$i] = 1 }
    foreach ($t in @($meta.text_cols)) { $types[[int]$t - 1] = 2 }
    # several background QueryTables at once (one per row block of the dump)
    $qts = @()
    foreach ($ch in @($meta.chunks)) {
        $qt = $d.QueryTables.Add("TEXT;" + $ch.file, $d.Cells([int]$ch.row, 1))
        $qt.TextFileParseType = 1
        $qt.TextFileCommaDelimiter = $false
        $qt.TextFileTabDelimiter = $true
        $qt.TextFileTextQualifier = -4142   # no quote handling: quotes/commas in data stay literal
        $qt.TextFilePlatform = 65001
        $qt.TextFileColumnDataTypes = $types
        $qt.AdjustColumnWidth = $false
        $qt.BackgroundQuery = $true
        $null = $qt.Refresh($true)
        $qts += $qt
    }
    while (@($qts | Where-Object { $_.Refreshing }).Count -gt 0) { Start-Sleep -Milliseconds 500 }
    foreach ($qt in $qts) { $qt.Delete() }
    $d.Name = "Sent Cases (CTS) - Bridge"
    $d.Rows(1).Font.Bold = $true
    T 'csv imported'

    $ncols = $d.UsedRange.Columns.Count
    $headers = @{}
    for ($c = 1; $c -le $ncols; $c++) { $headers[[string]$d.Cells(1, $c).Value2] = $c }
    function Col-Letter([int]$n) { $s = ''; while ($n -gt 0) { $r = ($n - 1) % 26; $s = [char](65 + $r) + $s; $n = [math]::Floor(($n - 1) / 26) }; $s }
    function L($name) { $l = Col-Letter $headers[$name]; return "'Sent Cases (CTS) - Bridge'!`$${l}:`$${l}" }
    foreach ($n in $meta.date_cols) { $d.Columns($headers[$n]).NumberFormat = 'dd-mmm-yy' }
    foreach ($n in $meta.datetime_cols) { if ($headers.ContainsKey($n)) { $d.Columns($headers[$n]).NumberFormat = 'dd-mmm-yy hh:mm:ss' } }
    # widths from the first rows only (cheap), so dates never show as ########
    $null = $d.Range($d.Cells(1, 1), $d.Cells(200, $ncols)).Columns.AutoFit()
    for ($c = 1; $c -le $ncols; $c++) {
        if ($d.Columns($c).ColumnWidth -gt 45) { $d.Columns($c).ColumnWidth = 45 }
        if ($d.Columns($c).ColumnWidth -lt 12) { $d.Columns($c).ColumnWidth = 12 }
    }

    foreach ($n in "Sent Summary", "Client Wise Check Sent", "Cluster Wise Check Sent", "Manual sentcount") {
        $s = $wb.Worksheets.Add($d); $s.Name = $n
    }
    $wb.Worksheets.Item("Sent Summary").Range("A1").Value2 = "Cluster Wise Summary"
    $wb.Worksheets.Item("Sent Summary").Range("A18").Value2 = "ReportType Summary"
    $wb.Worksheets.Item("Client Wise Check Sent").Range("A1").Value2 = "Bridge check Count Summary"
    $wb.Worksheets.Item("Cluster Wise Check Sent").Range("A2").Value2 = "Cluster Wise"

    # ---- Manual sentcount: B1 drop-down + SUMIFS grid (same shape as the reference)
    $ms = $wb.Worksheets.Item("Manual sentcount")
    $ms.Range("B1").Value2 = $meta.default_cluster
    $v = $ms.Range("B1").Validation
    $v.Delete()
    $v.Add(3, 1, 1, ($meta.clusters -join ','))
    $v.InCellDropdown = $true
    $ms.Range("B1").Font.Bold = $true
    $ms.Range("B1").Interior.Color = 13431551
    $ms.Range("D1").Value2 = "Manual Sent Count cluster wise"
    $ms.Range("D1").Font.Bold = $true
    $ms.Range("B2").Value2 = "Type"
    $ndays = [int]$meta.ndays
    $ms.Range("C2").Value2 = [double]$meta.start_serial
    if ($ndays -gt 1) { $ms.Range($ms.Cells(2, 4), $ms.Cells(2, 2 + $ndays)).Formula = "=C2+1" }
    $ms.Range($ms.Cells(2, 3), $ms.Cells(2, 2 + $ndays)).NumberFormat = 'dd-mmm-yy'
    $ms.Rows(2).Font.Bold = $true
    $rows = @(
        @("Address", "Address check count"), @("Database", "Database check count"),
        @("Education", "Education check count"), @("Employment", "Employment check count"))
    $dateRng = L 'LastDate'; $clRng = L 'Cluster'; $tyRng = L 'Type'
    $r = 3
    foreach ($row in $rows) {
        $ms.Cells($r, 2).Value2 = $row[0]
        $sumRng = L $row[1]
        $f = "=SUMIFS($sumRng,$dateRng,C`$2,$clRng,`$B`$1,$tyRng,""Manual"")"
        $ms.Range($ms.Cells($r, 3), $ms.Cells($r, 2 + $ndays)).Formula = $f
        $r++
    }
    $ms.Columns(2).ColumnWidth = 18
    T 'manual sentcount built'

    $ur = $d.UsedRange
    $src = $d.Range($d.Cells(1, 1), $d.Cells($ur.Rows.Count, $ur.Columns.Count))
    $pc = $wb.PivotCaches().Create(1, $src)
    function New-Pivot($sheet, $cell, $name) {
        $ws = $wb.Worksheets.Item($sheet)
        return $pc.CreatePivotTable($ws.Range($cell), $name)
    }
    function Set-Field($pt, $name, $orient, $pos) { $f = $pt.PivotFields($name); $f.Orientation = $orient; $f.Position = $pos }
    function Add-Data($pt, $field, $caption, $fn) {
        $df = $pt.AddDataField($pt.PivotFields($field), $caption, $fn); $df.NumberFormat = '#,##0'
    }
    $xlRow = 1; $xlCol = 2; $xlSum = -4157; $xlCount = -4112

    $p = New-Pivot "Sent Summary" "A3" "PivotTable2"
    Set-Field $p "Cluster" $xlRow 1; Set-Field $p "Report Status" $xlCol 1; Set-Field $p "Type" $xlCol 2
    Add-Data $p "report_sent_on" "Count of Cluster" $xlCount
    $p.PivotFields("Cluster").AutoSort(2, "Count of Cluster")

    $p = New-Pivot "Sent Summary" "A20" "PivotTable4"
    Set-Field $p "Cluster" $xlRow 1; Set-Field $p "Report Status" $xlCol 1; Set-Field $p "Report Type" $xlCol 2
    Add-Data $p "report_sent_on" "Count of Cluster" $xlCount
    $p.PivotFields("Cluster").AutoSort(2, "Count of Cluster")

    $p = New-Pivot "Client Wise Check Sent" "A3" "PivotTable1"
    Set-Field $p "Client Code" $xlRow 1; Set-Field $p "Client_name" $xlRow 2
    Add-Data $p "Report Checks Count" "Sum of Report Checks Count" $xlSum
    $p.RowAxisLayout(1)
    foreach ($n in "Client Code", "Client_name") {
        try { $p.PivotFields($n).Subtotals = @($false,$false,$false,$false,$false,$false,$false,$false,$false,$false,$false,$false) } catch {}
    }
    $p.PivotFields("Client Code").AutoSort(2, "Sum of Report Checks Count")

    $p = New-Pivot "Cluster Wise Check Sent" "A4" "PivotTable1"
    Set-Field $p "Cluster" $xlRow 1
    foreach ($c in "Report Checks Count", "Database check count", "Address check count", "Education check count", "Employment check count") {
        Add-Data $p $c "Sum of $c" $xlSum
    }
    T 'pivots built'

    $wb.Worksheets.Item("Sent Summary").Activate()
    $x.Calculation = -4105
    $x.Calculate()
    $wb.SaveAs($out, $fmt)
    T 'saved'
    $wb.Close($false)
} finally {
    $x.Quit()
}
