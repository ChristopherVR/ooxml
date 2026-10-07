param([string]$OutputFolder = (Join-Path $env:TEMP 'ooxml-native-chart-title-layout'))
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force -Path $OutputFolder | Out-Null
$excel = $null; $book = $null; $paths = @()
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false
    $book = $excel.Workbooks.Add(); $sheet = $book.Worksheets.Item(1)
    $sheet.Cells.Item(1,1).Value2 = [string]'Month'; $sheet.Cells.Item(1,2).Value2 = [string]'Sales'
    $sheet.Cells.Item(2,1).Value2 = [string]'A'; $sheet.Cells.Item(2,2).Value2 = [double]10
    $sheet.Cells.Item(3,1).Value2 = [string]'B'; $sheet.Cells.Item(3,2).Value2 = [double]20
    foreach ($name in @('automatic','automatic-overlay','moved-left','moved-right','moved-overlay','moved-long','moved-small')) {
        $width = if ($name -eq 'moved-small') { 320 } else { 480 }
        $object = $sheet.ChartObjects().Add(20,20,$width,300); $chart = $object.Chart
        $chart.ChartType = 51; $chart.SetSourceData($sheet.Range('A1:B3'),2)
        $chart.HasTitle = $true; $chart.HasLegend = $true
        $chart.ChartTitle.Text = if ($name -eq 'moved-long') { 'Revenue growth and forecast for the entire international business region' } else { 'Revenue' }
        $chart.ChartTitle.Font.Size = 18; $chart.ChartTitle.Font.Bold = $false
        $chart.ChartTitle.Font.Italic = $false; $chart.ChartTitle.Font.Name = 'Arial'; $chart.ChartTitle.Font.Color = 255
        $chart.ChartTitle.IncludeInLayout = $name -ne 'moved-overlay' -and $name -ne 'automatic-overlay'
        if (-not $name.StartsWith('automatic')) {
            $chart.ChartTitle.Left = if ($name -eq 'moved-right') { 340 } else { 40 }
            $chart.ChartTitle.Top = if ($name -eq 'moved-overlay') { 120 } else { 60 }
        }
        $chart.Parent.Activate(); $chart.Refresh()
        $path = Join-Path $OutputFolder "layout-$name.xlsx"
        $book.SaveCopyAs($path); $paths += $path
        if (-not $chart.Export((Join-Path $OutputFolder "layout-$name.png"),'PNG')) { throw 'Native chart export failed' }
        $object.Delete()
    }
} finally {
    if ($null -ne $book) { $book.Close($false); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($book) }
    if ($null -ne $excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
}
& (Join-Path $PSScriptRoot 'record-xlsx-chart-styles.ps1') -OutputFolder $OutputFolder -ReferenceFiles $paths -CaptureTitleCharacters -CaptureLayoutGeometry
