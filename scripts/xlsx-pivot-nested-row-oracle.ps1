# Manual native audit of two real row fields followed by Values. Outputs remain temporary.
$ErrorActionPreference = 'Stop'
$taskRoot = Split-Path -Parent $PSScriptRoot
$taskOutput = Join-Path ([IO.Path]::GetTempPath()) ('ooxml-nested-row-values-oracle-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $taskOutput | Out-Null
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false; $excel.DisplayAlerts = $false; $book = $null
try {
    $version = [Diagnostics.FileVersionInfo]::GetVersionInfo((Join-Path $excel.Path 'EXCEL.EXE')).FileVersion
    foreach ($layout in @('tabular','compact','outline')) {
        foreach ($column in @($false,$true)) {
            foreach ($totals in @('top','bottom','none')) {
                $profile = $layout + '-' + $totals + $(if ($column) { '-columns' } else { '' })
                $book = $excel.Workbooks.Open((Join-Path $taskRoot 'src/core/xlsx/__fixtures__/excel-pivot-values-rows-columns.xlsx'))
                $sheet = $book.Worksheets.Item('Summary'); $pivot = $sheet.PivotTables().Item(1)
                if (!$column) { $pivot.PivotFields('Year').Orientation = 0 }
                $pivot.DataPivotField.Orientation = 1; $pivot.DataPivotField.Position = 3
                $pivot.RowAxisLayout($(switch ($layout) { 'compact' { 0 } 'outline' { 2 } default { 1 } }))
                if ($totals -eq 'none') {
                    $pivot.PivotFields('Category').Subtotals = @($false,$false,$false,$false,$false,$false,$false,$false,$false,$false,$false,$false)
                } else {
                    $pivot.PivotFields('Category').LayoutSubtotalLocation = $(if ($totals -eq 'top') { 1 } else { 2 })
                }
                [void]$pivot.RefreshTable()
                $rows = @(); $indents = @()
                for ($r=1; $r -le $pivot.TableRange2.Rows.Count; $r++) {
                    $line = @()
                    for ($c=1; $c -le $pivot.TableRange2.Columns.Count; $c++) { $line += ,$sheet.Cells.Item($r,$c).Value2 }
                    $rows += ,$line; $indents += ,$sheet.Cells.Item($r,1).IndentLevel
                }
                $json = @{ values=$rows; indents=$indents; rowFields=$pivot.RowFields().Count; excelFileVersion=$version } | ConvertTo-Json -Depth 8
                [IO.File]::WriteAllText((Join-Path $taskOutput ($profile+'.json')), $json, [Text.UTF8Encoding]::new($false))
                $book.SaveAs((Join-Path $taskOutput ($profile+'.xlsx')),51)
                $book.Close($false); $book = $null
                Write-Output ($profile + ': ' + $rows.Count + ' rows, ' + $rows[0].Count + ' columns')
            }
        }
    }
} finally {
    if ($book) { $book.Close($false) }
    $excel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
}
Write-Output ('Oracle folder: ' + $taskOutput)
