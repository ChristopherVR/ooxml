param([string]$OutputFolder = (Join-Path $env:TEMP 'ooxml-native-chart-title-spacing'))
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
    $object = $sheet.ChartObjects().Add(20,20,480,300); $chart = $object.Chart
    $chart.ChartType = 51; $chart.SetSourceData($sheet.Range('A1:B3'),2)
    $chart.HasTitle = $true; $chart.HasLegend = $true
    foreach ($name in @('default','percent-150','percent-75','points-30','before-6','after-8','before-after','before-percent-50','after-percent-25','paragraphs','single-before-after','single-percent-150')) {
        $chart.ChartTitle.Text = if ($name.StartsWith('single')) { 'Revenue' } elseif ($name -eq 'paragraphs') { "Revenue`rForecast" } else { "Revenue`nForecast" }
        $chart.ChartTitle.Font.Size = 18; $chart.ChartTitle.Font.Bold = $false
        $chart.ChartTitle.Font.Italic = $false; $chart.ChartTitle.Font.Name = 'Arial'
        $chart.ChartTitle.Font.Color = 255
        if (-not $name.StartsWith('single')) { $chart.ChartTitle.Characters(9,8).Font.Color = 32768 }
        $paragraph = $chart.ChartTitle.Format.TextFrame2.TextRange.ParagraphFormat
        $paragraph.LineRuleWithin = -1; $paragraph.SpaceWithin = 1
        $paragraph.LineRuleBefore = 0; $paragraph.SpaceBefore = 0
        $paragraph.LineRuleAfter = 0; $paragraph.SpaceAfter = 0
        switch ($name) {
            'percent-150' { $paragraph.SpaceWithin = 1.5 }
            'percent-75' { $paragraph.SpaceWithin = 0.75 }
            'points-30' { $paragraph.LineRuleWithin = 0; $paragraph.SpaceWithin = 30 }
            'before-6' { $paragraph.SpaceBefore = 6 }
            'after-8' { $paragraph.SpaceAfter = 8 }
            'before-after' { $paragraph.SpaceBefore = 6; $paragraph.SpaceAfter = 8 }
            'before-percent-50' { $paragraph.LineRuleBefore = -1; $paragraph.SpaceBefore = 0.5 }
            'after-percent-25' { $paragraph.LineRuleAfter = -1; $paragraph.SpaceAfter = 0.25 }
            'paragraphs' { $paragraph.SpaceBefore = 6; $paragraph.SpaceAfter = 8 }
            'single-before-after' { $paragraph.SpaceBefore = 6; $paragraph.SpaceAfter = 8 }
            'single-percent-150' { $paragraph.SpaceWithin = 1.5 }
        }
        $chart.Parent.Activate(); $chart.Refresh()
        $path = Join-Path $OutputFolder "spacing-$name.xlsx"
        $book.SaveCopyAs($path); $paths += $path
        if (-not $chart.Export((Join-Path $OutputFolder "spacing-$name.png"),'PNG')) { throw 'Native chart export failed' }
    }
} finally {
    if ($null -ne $book) { $book.Close($false); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($book) }
    if ($null -ne $excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
}
& (Join-Path $PSScriptRoot 'record-xlsx-chart-styles.ps1') -OutputFolder $OutputFolder -ReferenceFiles $paths -CaptureTitleCharacters
