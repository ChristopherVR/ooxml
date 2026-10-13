param([Parameter(Mandatory=$true)][string]$ArtifactDirectory)
$ErrorActionPreference = 'Stop'
$nativeDir = (Get-Item -LiteralPath $ArtifactDirectory).FullName
if (Get-Process POWERPNT,EXCEL -ErrorAction SilentlyContinue) { throw 'Close existing PowerPoint and Excel processes before this isolated check.' }
$app = $null
$results = [System.Collections.Generic.List[object]]::new()
try {
    $app = New-Object -ComObject PowerPoint.Application
    foreach ($name in @('source-labels', 'labels-replaced', 'labels-resaved', 'native-resaved')) {
        $presentation = $null; $shape = $null; $chart = $null; $workbook = $null; $excel = $null
        try {
            $presentation = $app.Presentations.Open((Get-Item -LiteralPath (Join-Path $nativeDir "$name.pptx")).FullName, -1, 0, 0)
            $shape = $presentation.Slides.Item(1).Shapes.Item(1)
            $chart = $shape.Chart
            $chart.ChartData.Activate()
            $workbook = $chart.ChartData.Workbook
            $excel = $workbook.Application
            $excel.Visible = $false
            $sheet = $workbook.Worksheets.Item(1)
            $expectedName = if ($name -eq 'source-labels') { 'Dataset' } else { 'Renamed dataset' }
            $expectedCategory = if ($name -eq 'source-labels') { 'Q1' } else { 'First quarter' }
            if ($sheet.Range('B1').Value2 -ne $expectedName -or $sheet.Range('A2').Value2 -ne $expectedCategory) { throw "$name workbook labels disagree" }
            if ($sheet.Range('B2').Value2 -ne 10 -or $sheet.Range('B3').Value2 -ne 20 -or $sheet.Range('Z20').Value2 -ne '001 untouched ') { throw "$name changed unrelated workbook values" }
            $chart.Refresh()
            $series = $chart.SeriesCollection(1)
            if ($series.Name -ne $expectedName -or (@($series.XValues) -join ',') -ne "$expectedCategory,Q2" -or (@($series.Values) -join ',') -ne '10,20') { throw "$name chart caches disagree after native refresh" }
            if ($name -eq 'labels-resaved') { $presentation.SaveCopyAs((Join-Path $nativeDir 'native-resaved.pptx'), 24) }
            $results.Add(@{ file = "$name.pptx"; seriesName = $series.Name; categories = @($series.XValues); workbookName = $sheet.Range('B1').Value2; workbookCategory = $sheet.Range('A2').Value2; unchangedCell = $sheet.Range('Z20').Value2 })
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($series)
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($sheet)
        } finally {
            if ($workbook) { $workbook.Close($false); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) }
            if ($excel) { if ($excel.Workbooks.Count -eq 0) { $excel.Quit() }; [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
            foreach ($object in @($chart, $shape)) { if ($object) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($object) } }
            if ($presentation) { $presentation.Close(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($presentation) }
        }
    }
    @{ checkedAtUtc = [DateTime]::UtcNow.ToString('o'); officeVersion = $app.Version; cases = $results } | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $nativeDir 'chart-label-acceptance.json')
    Write-Output "Passed $($results.Count) native chart-label workbook activation, refresh and reopen checks."
} finally {
    if ($app) { if ($app.Presentations.Count -eq 0) { $app.Quit() }; [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app) }
}
