param(
    [string]$OutputFolder = (Join-Path $env:TEMP 'ooxml-native-chart-styles'),
    [int[]]$StyleIds = (201..216),
    [switch]$BuiltInReferences,
    [string[]]$ReferenceFiles = @(),
    [switch]$CaptureTitleCharacters,
    [switch]$CaptureLayoutGeometry
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
New-Item -ItemType Directory -Force -Path $OutputFolder | Out-Null
$excel = $null; $book = $null
if ($BuiltInReferences -and -not $PSBoundParameters.ContainsKey('StyleIds')) {
    $StyleIds = (1..48) + (101..148)
}
function Read-ChartFont($font, [switch]$TextRange) {
    $size = [double]$font.Size
    $rgb = if ($TextRange) { [int]$font.Fill.ForeColor.RGB } else { [int]$font.Color }
    return @{ size = $(if ($size -gt 0) { $size } else { $null }); name = [string]$font.Name; bold = [bool]$font.Bold; italic = [bool]$font.Italic; color = '#{0:X2}{1:X2}{2:X2}' -f ($rgb -band 255),(($rgb -shr 8) -band 255),(($rgb -shr 16) -band 255) }
}
function Read-ChartPaint($format) {
    $fillRgb = [int]$format.Fill.ForeColor.RGB
    $lineRgb = [int]$format.Line.ForeColor.RGB
    $weight = [double]$format.Line.Weight
    return @{ fillColor = '#{0:X2}{1:X2}{2:X2}' -f ($fillRgb -band 255),(($fillRgb -shr 8) -band 255),(($fillRgb -shr 16) -band 255); lineColor = $(if ($weight -gt 0) { '#{0:X2}{1:X2}{2:X2}' -f ($lineRgb -band 255),(($lineRgb -shr 8) -band 255),(($lineRgb -shr 16) -band 255) } else { $null }); lineWeight = $(if ($weight -gt 0) { $weight } else { $null }) }
}
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false
    $book = $excel.Workbooks.Add(); $sheet = $book.Worksheets.Item(1)
    $sheet.Cells.Item(1,1).Value2 = [string]'Month'
    $sheet.Cells.Item(1,2).Value2 = [string]'Sales'
    $sheet.Cells.Item(1,3).Value2 = [string]'Cost'
    for ($row=2; $row -le 5; $row++) {
        $sheet.Cells.Item($row,1).Value2 = [string]"M$row"
        $sheet.Cells.Item($row,2).Value2 = [double]($row*2)
        $sheet.Cells.Item($row,3).Value2 = [double]$row
    }
    $scheme = @{}
    $names = @('dk1','lt1','dk2','lt2','accent1','accent2','accent3','accent4','accent5','accent6','hlink','folHlink')
    for ($i=0; $i -lt $names.Count; $i++) { $rgb=[int]$book.Theme.ThemeColorScheme.Colors($i+1).RGB; $scheme[$names[$i]] = '#{0:X2}{1:X2}{2:X2}' -f ($rgb -band 255),(($rgb -shr 8) -band 255),(($rgb -shr 16) -band 255) }
    $cases = @()
    $references = if ($ReferenceFiles.Count) {
        @($ReferenceFiles | ForEach-Object { @{ path = (Resolve-Path -LiteralPath $_).Path } })
    } else { @($StyleIds | ForEach-Object { @{ id = $_ } }) }
    foreach ($reference in $references) {
        $id = $reference.id
        if ($reference.path) {
            $path = $reference.path
        } elseif ($BuiltInReferences) {
            $path = Join-Path $OutputFolder "builtin-$id.xlsx"
        } else {
            $object = $sheet.ChartObjects().Add(260,20,480,300)
            $chart = $object.Chart; $chart.ChartType = 51
            $chart.SetSourceData($sheet.Range('A1:C5'),2)
            $chart.ChartStyle = $id; $chart.ChartColor = 10
            $chart.HasTitle = $true; $chart.ChartTitle.Text = 'Native style'
            $path = Join-Path $OutputFolder "column-$id.xlsx"
            $book.SaveCopyAs($path)
        }
        $zip = [System.IO.Compression.ZipFile]::OpenRead($path)
        try {
            $parts = @{}
            foreach ($entry in $zip.Entries) {
                if ($entry.FullName -notmatch '^xl/charts/') { continue }
                $reader = [System.IO.StreamReader]::new($entry.Open())
                try { $parts[$entry.FullName] = $reader.ReadToEnd() } finally { $reader.Dispose() }
            }
        } finally { $zip.Dispose() }
        $probeBook = $excel.Workbooks.Open($path,0,$true)
        try {
            $probeChart = $probeBook.Worksheets.Item(1).ChartObjects(1).Chart
            $probeChart.Parent.Activate(); $probeChart.Refresh()
            $cases += @{ requestedStyle = $id; style = [int]$probeChart.ChartStyle; titleFontSize = $(if ([double]$probeChart.ChartTitle.Font.Size -gt 0) { [double]$probeChart.ChartTitle.Font.Size } else { $null }); axisFontSize = [double]$probeChart.Axes(1).TickLabels.Font.Size; legendFontSize = [double]$probeChart.Legend.Font.Size; titleText = (Read-ChartFont $probeChart.ChartTitle.Font); categoryText = (Read-ChartFont $probeChart.Axes(1).TickLabels.Font); valueText = (Read-ChartFont $probeChart.Axes(2).TickLabels.Font); legendText = (Read-ChartFont $probeChart.Legend.Font); parts = $parts }
            $cases[-1].hasValueAxis = [bool]$probeChart.HasAxis(2,1)
            $cases[-1].chartArea = Read-ChartPaint $probeChart.ChartArea.Format
            if ($reference.path) {
                $cases[-1].Remove('requestedStyle')
                $cases[-1].referenceName = [System.IO.Path]::GetFileNameWithoutExtension($path)
                # The title-wide Font can differ from the actual rich-text range.
                $cases[-1].titleText = Read-ChartFont $probeChart.ChartTitle.Format.TextFrame2.TextRange.Font -TextRange
            }
            if ($CaptureLayoutGeometry) {
                $cases[-1].layoutGeometry = @{
                    chart = @{ widthPt = [double]$probeChart.Parent.Width; heightPt = [double]$probeChart.Parent.Height }
                    chartArea = @{ leftPt = [double]$probeChart.ChartArea.Left; topPt = [double]$probeChart.ChartArea.Top; widthPt = [double]$probeChart.ChartArea.Width; heightPt = [double]$probeChart.ChartArea.Height }
                    plotArea = @{ leftPt = [double]$probeChart.PlotArea.Left; topPt = [double]$probeChart.PlotArea.Top; widthPt = [double]$probeChart.PlotArea.Width; heightPt = [double]$probeChart.PlotArea.Height }
                    insidePlotArea = @{ leftPt = [double]$probeChart.PlotArea.InsideLeft; topPt = [double]$probeChart.PlotArea.InsideTop; widthPt = [double]$probeChart.PlotArea.InsideWidth; heightPt = [double]$probeChart.PlotArea.InsideHeight }
                    titleIncludeInLayout = [bool]$probeChart.ChartTitle.IncludeInLayout
                }
            }
            if ($CaptureTitleCharacters) {
				$cases[-1].titleGeometry = @{
					leftPt = [double]$probeChart.ChartTitle.Left; topPt = [double]$probeChart.ChartTitle.Top
					widthPt = [double]$probeChart.ChartTitle.Width; heightPt = [double]$probeChart.ChartTitle.Height
				}
                $title = [string]$probeChart.ChartTitle.Text
                $cases[-1].titleCharacters = @(for ($index = 0; $index -lt $title.Length; $index++) {
                    $font = $probeChart.ChartTitle.Characters($index+1,1).Font
                    $properties = Read-ChartFont $font
                    $properties.underline = [int]$font.Underline -ne -4142
                    @{ text = $title.Substring($index,1); font = $properties }
                })
            }
        } finally { $probeBook.Close($false) }
        if (-not $BuiltInReferences -and -not $reference.path) { $object.Delete() }
    }
    $outputName = if ($ReferenceFiles.Count) { 'references.json' } elseif ($BuiltInReferences) { 'built-in-styles.json' } else { 'styles.json' }
    @{ excelVersion = [string]$excel.Version; excelBuild = [string]$excel.Build; scheme = $scheme; cases = $cases } |
        ConvertTo-Json -Depth 8 | Set-Content -Encoding utf8 (Join-Path $OutputFolder $outputName)
    Write-Output "Recorded $($cases.Count) independent native style parts in $OutputFolder"
} finally {
    if ($null -ne $book) { $book.Close($false) }
    if ($null -ne $excel) { $excel.Quit() }
}
