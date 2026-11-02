$ErrorActionPreference='Stop'
$src=$args[0]; $out=$args[1]
$x=New-Object -ComObject Excel.Application; $x.Visible=$false; $x.DisplayAlerts=$false
$null=$x.Workbooks.Add(); $x.Calculation=-4135
$lines=@()
try{
  $w=$x.Workbooks.Open($src,0,$true)
  foreach($ws in $w.Worksheets){
    $lines+="SHEET $($ws.Name) pivots=$($ws.PivotTables().Count)"
    foreach($pt in $ws.PivotTables()){
      $lines+="  PT $($pt.Name) at $($pt.TableRange2.Address()) src=$($pt.SourceData) cache=$($pt.CacheIndex) layout=$($pt.RowAxisLayout(1)) compact=$($pt.CompactLayoutRowHeader)"
      foreach($pf in $pt.PivotFields()){
        if($pf.Orientation -ne 0){
          $lines+="    field '$($pf.Name)' orient=$($pf.Orientation) pos=$($pf.Position) fn=$($pf.Function) caption=$($pf.Caption) subtotals1=$($pf.Subtotals(1))"
          if($pf.Orientation -eq 3){ try{$lines+="      page current=$($pf.CurrentPage.Name) multi=$($pf.EnableMultiplePageItems)"}catch{} }
        }
      }
      foreach($df in $pt.DataFields){ $lines+="    DATA '$($df.Name)' src=$($df.SourceName) fn=$($df.Function) fmt=$($df.NumberFormat)" }
      $lines+="    grandRows=$($pt.RowGrand) grandCols=$($pt.ColumnGrand) style=$($pt.TableStyle2)"
    }
  }
  $w.Close($false)
} finally { $x.Quit(); $lines | Set-Content $out -Encoding UTF8 }
