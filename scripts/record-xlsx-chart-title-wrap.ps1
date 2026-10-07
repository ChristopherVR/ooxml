param([string]$OutputFolder = (Join-Path $env:TEMP 'ooxml-native-chart-title-wrap'))
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force -Path $OutputFolder | Out-Null
$excel = $null; $book = $null; $paths = @(); $dimensions = @{}
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
    foreach ($name in @('short-480','long-480','long-320','long-200','mixed-320','word-160','word-200','word-240','word-320','spaced-480')) {
        $width = [double]($name.Split('-')[-1]); $object.Width = $width
        $chart.ChartTitle.Text = if ($name.StartsWith('short')) { 'Revenue' } elseif ($name.StartsWith('word')) { 'InternationalBusinessRevenueForecast' } else { 'Revenue growth and forecast for the entire international business region' }
        $chart.ChartTitle.Font.Size = 18; $chart.ChartTitle.Font.Bold = $false
        $chart.ChartTitle.Font.Italic = $false; $chart.ChartTitle.Font.Name = 'Arial'
        $chart.ChartTitle.Font.Color = 255
        $paragraph = $chart.ChartTitle.Format.TextFrame2.TextRange.ParagraphFormat
        $paragraph.LineRuleBefore = 0; $paragraph.SpaceBefore = 0
        $paragraph.LineRuleAfter = 0; $paragraph.SpaceAfter = 0
        if ($name.StartsWith('spaced')) { $paragraph.SpaceBefore = 6; $paragraph.SpaceAfter = 8 }
        if ($name.StartsWith('mixed')) { $chart.ChartTitle.Characters(1,7).Font.Size = 24; $chart.ChartTitle.Characters(1,7).Font.Bold = $true }
        $chart.Parent.Activate(); $chart.Refresh()
        $path = Join-Path $OutputFolder "wrap-$name.xlsx"
        $book.SaveCopyAs($path); $paths += $path
        $dimensions["wrap-$name"] = @{ widthPt=$width; heightPt=[double]$object.Height }
        if (-not $chart.Export((Join-Path $OutputFolder "wrap-$name.png"),'PNG')) { throw 'Native chart export failed' }
    }
} finally {
    if ($null -ne $book) { $book.Close($false); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($book) }
    if ($null -ne $excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
}
& (Join-Path $PSScriptRoot 'record-xlsx-chart-styles.ps1') -OutputFolder $OutputFolder -ReferenceFiles $paths -CaptureTitleCharacters
$records = Get-Content (Join-Path $OutputFolder 'references.json') -Raw | ConvertFrom-Json
foreach ($case in $records.cases) { $case | Add-Member -NotePropertyName chartGeometry -NotePropertyValue $dimensions[$case.referenceName] }
$records | ConvertTo-Json -Depth 9 | Set-Content -Encoding utf8 (Join-Path $OutputFolder 'references.json')
