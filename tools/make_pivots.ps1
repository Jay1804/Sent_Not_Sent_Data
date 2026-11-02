# usage: make_pivots.ps1 <in.xlsx> <out> <fileformat 50=xlsb|51=xlsx>
$ErrorActionPreference = 'Stop'
$in = $args[0]; $out = $args[1]; $fmt = [int]$args[2]
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
try {
    $null = $x.Workbooks.Add(); $x.Calculation = -4135
    $wb = $x.Workbooks.Open($in)
    $d = $wb.Worksheets.Item("Sent Cases (CTS) - Bridge")
    $ur = $d.UsedRange
    $src = $d.Range($d.Cells(1, 1), $d.Cells($ur.Rows.Count, $ur.Columns.Count))
    $pc = $wb.PivotCaches().Create(1, $src)

    function New-Pivot($sheet, $cell, $name) {
        $ws = $wb.Worksheets.Item($sheet)
        return $pc.CreatePivotTable($ws.Range($cell), $name)
    }
    function Set-Field($pt, $name, $orient, $pos) {
        $f = $pt.PivotFields($name); $f.Orientation = $orient; $f.Position = $pos
    }
    function Add-Data($pt, $field, $caption, $fn) {
        $df = $pt.AddDataField($pt.PivotFields($field), $caption, $fn)
        $df.NumberFormat = '#,##0'
    }
    $xlRow = 1; $xlCol = 2; $xlSum = -4157; $xlCount = -4112

    # Sent Summary: Cluster x (Report Status > Type), then Cluster x (Report Status > Report Type)
    $p = New-Pivot "Sent Summary" "A3" "PivotTable2"
    Set-Field $p "Cluster" $xlRow 1; Set-Field $p "Report Status" $xlCol 1; Set-Field $p "Type" $xlCol 2
    Add-Data $p "report_sent_on" "Count of Cluster" $xlCount
    $p.PivotFields("Cluster").AutoSort(2, "Count of Cluster")

    $p = New-Pivot "Sent Summary" "A20" "PivotTable4"
    Set-Field $p "Cluster" $xlRow 1; Set-Field $p "Report Status" $xlCol 1; Set-Field $p "Report Type" $xlCol 2
    Add-Data $p "report_sent_on" "Count of Cluster" $xlCount
    $p.PivotFields("Cluster").AutoSort(2, "Count of Cluster")

    # Client Wise Check Sent: Client Code > Client_name, sum of report checks (tabular, no subtotals)
    $p = New-Pivot "Client Wise Check Sent" "A3" "PivotTable1"
    Set-Field $p "Client Code" $xlRow 1; Set-Field $p "Client_name" $xlRow 2
    Add-Data $p "Report Checks Count" "Sum of Report Checks Count" $xlSum
    $p.RowAxisLayout(1)
    foreach ($n in "Client Code", "Client_name") {
        try { $p.PivotFields($n).Subtotals = @($false,$false,$false,$false,$false,$false,$false,$false,$false,$false,$false,$false) } catch {}
    }
    $p.PivotFields("Client Code").AutoSort(2, "Sum of Report Checks Count")

    # Cluster Wise Check Sent
    $p = New-Pivot "Cluster Wise Check Sent" "A4" "PivotTable1"
    Set-Field $p "Cluster" $xlRow 1
    foreach ($c in "Report Checks Count", "Database check count", "Address check count", "Education check count", "Employment check count") {
        Add-Data $p $c "Sum of $c" $xlSum
    }

    $x.Calculation = -4105
    $wb.SaveAs($out, $fmt)
    $wb.Close($false)
} finally {
    $x.Quit()
}
