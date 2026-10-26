# Native PowerPoint acceptance for repeated classic saves and the embedded Edit Data workbook.
# Run manually on Windows with PowerPoint/Excel installed and no existing Office sessions.
param([Parameter(Mandatory=$true)][string]$OutDir, [switch]$SkipBuild, [switch]$RequireActiveDataSheet)
$ErrorActionPreference = 'Stop'
$OutDir = [IO.Path]::GetFullPath($OutDir)
if (Get-Process POWERPNT,EXCEL -ErrorAction SilentlyContinue) { throw 'Close existing PowerPoint and Excel sessions before this private acceptance run.' }
if (-not $SkipBuild) {
 & bun (Join-Path $PSScriptRoot 'pptx-classic-chart-acceptance-build.ts') $OutDir
 if ($LASTEXITCODE -ne 0) { throw 'Classic chart acceptance builder failed.' }
}
$manifest = Get-Content -Raw (Join-Path $OutDir 'manifest.json') | ConvertFrom-Json
$app = New-Object -ComObject PowerPoint.Application
$app.DisplayAlerts = 2
$results = @()
$failures = @()
function Assert-Same($actual, $expected, [string]$where) {
 $same = [string]$actual -eq [string]$expected
 if ($expected -is [ValueType]) {
  $number = 0.0
  $same = [double]::TryParse([string]$actual, [ref]$number) -and [math]::Abs($number-[double]$expected) -lt 0.00000001
 }
 if (-not $same) { throw "$where is '$actual', expected '$expected'." }
}
try {
 foreach ($property in $manifest.PSObject.Properties) {
  $presentation = $null
  $book = $null
  $excel = $null
  try {
   $want = $property.Value
   $presentation = $app.Presentations.Open((Join-Path $OutDir $property.Name), 0, 0, 0)
   Assert-Same $presentation.Slides.Count $want.slideCount "$($property.Name) slide count"
   $charts = @()
   foreach ($shape in $presentation.Slides.Item([int]$want.slideIndex).Shapes) { if ($shape.HasChart -eq -1) { $charts += $shape.Chart } }
   Assert-Same $charts.Count 1 "$($property.Name) chart count"
   $chart = $charts[0]
   Assert-Same $chart.ChartType $want.nativeType "$($property.Name) chart type"
   Assert-Same $chart.SeriesCollection().Count @($want.series).Count "$($property.Name) series count"
   for ($index=0; $index -lt @($want.series).Count; $index++) {
    $series = $chart.SeriesCollection($index+1)
    $expected = $want.series[$index]
    Assert-Same $series.Name $expected.name 'series name'
    $values = @($series.Values)
    Assert-Same $values.Count @($expected.values).Count 'series value count'
    for ($row=0; $row -lt $values.Count; $row++) { Assert-Same $values[$row] $expected.values[$row] 'series value' }
    $categories = @($series.XValues)
    Assert-Same $categories.Count @($want.categories).Count 'category count'
    for ($row=0; $row -lt $categories.Count; $row++) { Assert-Same $categories[$row] $want.categories[$row] 'category' }
   }
   $chart.ChartData.Activate()
   $book = $chart.ChartData.Workbook
   $excel = $book.Application
   $excel.Visible = $false
   $activeSheet = [string]$book.ActiveSheet.Name
   $referencedSheets = @($want.cells.sheet | Select-Object -Unique)
   if ($RequireActiveDataSheet -and $activeSheet -notin $referencedSheets) { throw "Edit Data opens '$activeSheet' instead of the referenced sheet(s): $($referencedSheets -join ', ')" }
   foreach ($cell in @($want.cells)) {
    $sheet = $book.Worksheets.Item([string]$cell.sheet)
    Assert-Same $sheet.Range([string]$cell.address).Value2 $cell.value "Edit Data $($cell.sheet)!$($cell.address)"
   }
   $results += [ordered]@{ file=$property.Name; version=[string]$app.Version; status='ok'; slides=$presentation.Slides.Count; charts=1; series=@($want.series).Count; cells=@($want.cells).Count; activeSheet=$activeSheet; referencedSheets=$referencedSheets }
  } catch {
   $failures += "$($property.Name): $($_.Exception.Message)"
   $results += [ordered]@{ file=$property.Name; status='failed'; detail=$_.Exception.Message }
  } finally {
   if ($null -ne $book) { try { $book.Close($false) } catch {} }
   if ($null -ne $excel) { try { $excel.Quit() } catch {} }
   if ($null -ne $presentation) { try { $presentation.Saved=-1; $presentation.Close() } catch {} }
  }
 }
} finally {
 try { $app.Quit() } catch {}
 [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)
 [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
$results | ConvertTo-Json -Depth 8 | Set-Content (Join-Path $OutDir 'results.json')
$results | Format-Table -AutoSize
if ($failures.Count) { throw ($failures -join "`n") }
Write-Host 'Native open, chart data and embedded workbook checks passed; this does not assert pixel parity or exclude every silent repair.'





