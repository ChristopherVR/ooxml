param([string]$OutputFolder = (Join-Path $env:TEMP 'ooxml-native-chart-title-text'))
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force -Path $OutputFolder | Out-Null
$excel = $null; $book = $null
$paths = @()
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false
    $book = $excel.Workbooks.Add(); $sheet = $book.Worksheets.Item(1)
    $sheet.Cells.Item(1,1).Value2 = [string]'Month'; $sheet.Cells.Item(1,2).Value2 = [string]'Sales'
    $sheet.Cells.Item(2,1).Value2 = [string]'A'; $sheet.Cells.Item(2,2).Value2 = [double]10
    $sheet.Cells.Item(3,1).Value2 = [string]'B'; $sheet.Cells.Item(3,2).Value2 = [double]20
    $object = $sheet.ChartObjects().Add(20,20,480,300); $chart = $object.Chart
    $chart.ChartType = 51; $chart.SetSourceData($sheet.Range('A1:B3'),2)
    $chart.HasTitle = $true; $chart.HasLegend = $true
    foreach ($name in @('single','lines')) {
        $chart.ChartTitle.Text = if ($name -eq 'single') { 'Revenue growth' } else { "Revenue growth`nForecast" }
        $chart.ChartTitle.Font.Size = 18; $chart.ChartTitle.Font.Bold = $false
        $chart.ChartTitle.Font.Italic = $false; $chart.ChartTitle.Font.Underline = -4142
        $chart.ChartTitle.Font.Color = 0; $chart.ChartTitle.Font.Name = 'Aptos Narrow'
        $revenue = $chart.ChartTitle.Characters(1,7).Font
        $revenue.Size = 24; $revenue.Bold = $true; $revenue.Color = 255; $revenue.Name = 'Arial'
        $growth = $chart.ChartTitle.Characters(9,6).Font
        $growth.Size = 12; $growth.Italic = $true; $growth.Color = 16711680
        if ($name -eq 'lines') {
            $forecast = $chart.ChartTitle.Characters(16,8).Font
            $forecast.Size = 20; $forecast.Name = 'Cambria'; $forecast.Color = 32768
            $forecast.Underline = 2
        }
        $chart.Parent.Activate(); $chart.Refresh()
        $path = Join-Path $OutputFolder "mixed-$name.xlsx"
        $book.SaveCopyAs($path); $paths += $path
        if (-not $chart.Export((Join-Path $OutputFolder "mixed-$name.png"),'PNG')) { throw 'Native chart export failed' }
    }
} finally {
    if ($null -ne $book) { $book.Close($false); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($book) }
    if ($null -ne $excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
}
& bun (Join-Path $PSScriptRoot 'prepare-xlsx-chart-title-paragraphs.ts') $OutputFolder
if ($LASTEXITCODE -ne 0) { throw 'Paragraph reference preparation failed' }
$paths += Join-Path $OutputFolder 'mixed-paragraphs.xlsx'
& (Join-Path $PSScriptRoot 'record-xlsx-chart-styles.ps1') -OutputFolder $OutputFolder -ReferenceFiles $paths -CaptureTitleCharacters
