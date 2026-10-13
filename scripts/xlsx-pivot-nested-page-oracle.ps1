param([switch]$RepeatLabels, [switch]$RepeatColumnLabels)
# Owned synthetic native reference, with output kept outside the repository.
$ErrorActionPreference = 'Stop'
$taskOutput = Join-Path ([IO.Path]::GetTempPath()) ('ooxml-nested-page-oracle-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory $taskOutput | Out-Null
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false; $excel.DisplayAlerts = $false; $book = $null
try {
    $version = [Diagnostics.FileVersionInfo]::GetVersionInfo((Join-Path $excel.Path 'EXCEL.EXE')).FileVersion
    foreach ($layout in @('tabular','compact','outline')) {
    foreach ($columns in @($false,$true)) {
    foreach ($selection in @('All','East','West')) {
    foreach ($subtotal in @('top','bottom','none')) {
        $name = "$layout-$selection-$subtotal" + $(if ($columns) {'-columns'} else {''}) + $(if ($RepeatLabels) {'-repeat'} else {''})
        $book = $excel.Workbooks.Add()
        $data = $book.Worksheets.Item(1); $data.Name = 'Data'
        $summary = $book.Worksheets.Add(); $summary.Name = 'Summary'
        $rows = @(@('Category','Region','Market','Year','Amount','Quantity'),
            @('Fruit','North','East',2025,10,2), @('Fruit','North','West',2026,20,3),
            @('Fruit','South','East',2026,7,4), @('Veg','South','West',2025,15,5),
            @('Veg','North','East',2025,5,6), @('Berry','Islands','West',2026,9,7),
            @('Fruit','South','East',2025,3,8))
        for ($r=0; $r -lt $rows.Count; $r++) { for ($c=0; $c -lt $rows[0].Count; $c++) {
            $v=$rows[$r][$c]
            if ($v -is [int]) { $data.Cells.Item($r+1,$c+1).Value2=[double]$v }
            else { $data.Cells.Item($r+1,$c+1).Value2=[string]$v }
        }}
        $pivot=$book.PivotCaches().Create(1,'Data!R1C1:R8C6',6).CreatePivotTable($summary.Range('A1'),'SalesPivot')
        $pivot.PivotFields('Category').Orientation=1; $pivot.PivotFields('Region').Orientation=1
        if ($columns) { $pivot.PivotFields('Year').Orientation=2 }
        $pivot.PivotFields('Market').Orientation=3
        [void]$pivot.AddDataField($pivot.PivotFields('Amount'),'Sum of Amount',-4157)
        [void]$pivot.AddDataField($pivot.PivotFields('Quantity'),'Count of Quantity',-4112)
        [void]$pivot.AddDataField($pivot.PivotFields('Amount'),'Average of Amount',-4106)
        $pivot.RowAxisLayout($(switch ($layout) {compact {0} outline {2} default {1}}))
        if ($subtotal -eq 'none') { $pivot.PivotFields('Category').Subtotals=@($false,$false,$false,$false,$false,$false,$false,$false,$false,$false,$false,$false) }
        else { $pivot.PivotFields('Category').LayoutSubtotalLocation=$(if ($subtotal -eq 'top') {1} else {2}) }
        if ($selection -ne 'All') { $pivot.PivotFields('Market').CurrentPage=$selection }
        if ($RepeatLabels) {
            if ($RepeatColumnLabels) { $pivot.RepeatAllLabels(2) }
            else {
                $pivot.PivotFields('Category').RepeatLabels=$true
                $pivot.PivotFields('Region').RepeatLabels=$true
            }
        }
        [void]$pivot.RefreshTable()
        $values=@(); $indents=@()
        for ($r=1; $r -le $pivot.TableRange2.Rows.Count; $r++) {
            $line=@(); for ($c=1; $c -le $pivot.TableRange2.Columns.Count; $c++) { $line+=,$summary.Cells.Item($r,$c).Value2 }
            $values+=,$line; $indents+=,$summary.Cells.Item($r,1).IndentLevel
        }
        [IO.File]::WriteAllText((Join-Path $taskOutput ($name+'.json')),(@{
            values=$values; indents=$indents; tableRange1=$pivot.TableRange1.Address();
            tableRange2=$pivot.TableRange2.Address(); excelFileVersion=$version;
            layout=$layout; columns=$columns; selection=$selection; subtotal=$subtotal; repeatLabels=[bool]$RepeatLabels; repeatColumnLabels=[bool]($RepeatLabels -and $RepeatColumnLabels)
        } | ConvertTo-Json -Depth 8),[Text.UTF8Encoding]::new($false))
        $book.SaveAs((Join-Path $taskOutput ($name+'.xlsx')),51); $book.Close($false); $book=$null
        Write-Output "$name : $($values.Count) rows"
    }}}}
} finally {
    if ($book) { $book.Close($false) }
    $excel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
}
Write-Output ('Oracle folder: '+$taskOutput)
