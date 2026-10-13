param([Parameter(Mandatory=$true)][string]$ArtifactDirectory)
$ErrorActionPreference = 'Stop'
$nativeDir = (Get-Item -LiteralPath $ArtifactDirectory).FullName
if (Get-Process POWERPNT -ErrorAction SilentlyContinue) { throw 'Close existing PowerPoint processes before this isolated check.' }
$app = $null
$results = [System.Collections.Generic.List[object]]::new()
try {
    $app = New-Object -ComObject PowerPoint.Application
    foreach ($name in @('source-rich', 'rich-replaced', 'rich-resaved', 'native-resaved', 'source-linked')) {
        $presentation = $null; $slide = $null; $shape = $null; $chart = $null; $title = $null; $series = $null
        try {
            $path = (Get-Item -LiteralPath (Join-Path $nativeDir "$name.pptx")).FullName
            $presentation = $app.Presentations.Open($path, -1, 0, 0)
            if ($presentation.Slides.Count -ne 1) { throw "$name has an unexpected slide count" }
            $slide = $presentation.Slides.Item(1)
            for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
                $candidate = $slide.Shapes.Item($i)
                if ($candidate.HasChart -eq -1) { $shape = $candidate; break }
                [void][Runtime.InteropServices.Marshal]::ReleaseComObject($candidate)
            }
            if (-not $shape) { throw "$name lost the chart" }
            $chart = $shape.Chart
            $title = $chart.ChartTitle
            $text = $title.Text.Trim()
            $expected = if ($name -in @('source-rich', 'source-linked')) { 'Annual report' } else { 'Annual summary' }
            if ($text -ne $expected) { throw "$name title is '$text', expected '$expected'" }
            $bold = [int]$title.Format.TextFrame2.TextRange.Characters(1, 6).Font.Bold
            if ($name -ne 'source-linked' -and $bold -ne -1) { throw "$name lost the starting run's bold format" }
            $series = $chart.SeriesCollection(1)
            $seriesName = $series.Name; $values = @($series.Values); $categories = @($series.XValues)
            if ($seriesName -ne 'Dataset' -or ($values -join ',') -ne '10,20' -or ($categories -join ',') -ne 'Q1,Q2') { throw "$name changed chart data while replacing title text" }
            if ($name -eq 'rich-resaved') { $presentation.SaveCopyAs((Join-Path $nativeDir 'native-resaved.pptx'), 24) }
            $results.Add(@{ file = "$name.pptx"; title = $text; firstRunBold = $bold; seriesName = $seriesName; values = $values; categories = $categories })
        } finally {
            foreach ($object in @($series, $title, $chart, $shape, $slide)) { if ($object) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($object) } }
            if ($presentation) { $presentation.Close(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($presentation) }
        }
    }
    @{ checkedAtUtc = [DateTime]::UtcNow.ToString('o'); officeVersion = $app.Version; cases = $results } | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $nativeDir 'chart-title-acceptance.json')
    Write-Output "Passed $($results.Count) native chart-title open/save/reopen checks."
} finally {
    if ($app) { if ($app.Presentations.Count -eq 0) { $app.Quit() }; [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app) }
}
