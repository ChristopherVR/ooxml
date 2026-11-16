# Manual native acceptance for the shared chart analysis and picture fill edits (not run in CI;
# needs desktop Word, Excel and PowerPoint on Windows, and no Office session of your own open).
#
#   pwsh -File scripts/chart-analysis-native-acceptance.ps1 -OutDir <folder> [-SkipBuild]
#
# `bun scripts/chart-analysis-native-build.ts` writes one Word document, one Excel workbook and two
# PowerPoint decks (see that file). Each is opened through its application's COM object model, one
# application at a time, with alerts off and Word's OpenAndRepair off. A file fails when it throws,
# when Excel writes a repair log or captions the window "Repaired", or when a chart differs from
# manifest.json: the series count, which series sit on the secondary axis (Series.AxisGroup), the
# trendline count and error bars of the analysed series, and whether the picture-filled series and
# plot area report a picture or texture fill (Fill.Type 4 or 6). This checks what the
# applications read back, not how they paint it.
param(
	[Parameter(Mandatory = $true)][string]$OutDir,
	[switch]$SkipBuild
)
$ErrorActionPreference = 'Stop'
$OutDir = [IO.Path]::GetFullPath($OutDir)
$root = Split-Path -Parent $PSScriptRoot
if (Get-Process WINWORD, EXCEL, POWERPNT -ErrorAction SilentlyContinue) {
	throw 'Close Word, Excel and PowerPoint before this acceptance run.'
}
if (-not $SkipBuild) {
	Push-Location $root
	try {
		& bun scripts/chart-analysis-native-build.ts $OutDir
		if ($LASTEXITCODE -ne 0) { throw "chart-analysis-native-build.ts failed ($LASTEXITCODE)" }
	} finally { Pop-Location }
}
$manifest = Get-Content -Raw (Join-Path $OutDir 'manifest.json') | ConvertFrom-Json
$pictureTypes = @(4, 6) # msoFillTextured, msoFillPicture

function Test-Chart($chart, $want, [string]$where) {
	$problems = @()
	$count = $chart.SeriesCollection().Count
	if ($count -ne [int]$want.series) { return @("$where has $count series, expected $($want.series)") }
	for ($s = 1; $s -le $count; $s++) {
		$group = [int]$chart.SeriesCollection($s).AxisGroup
		$expected = if (@($want.secondary) -contains $s) { 2 } else { 1 }
		if ($group -ne $expected) { $problems += "$where series $s axis group is $group, expected $expected" }
	}
	if ($want.trendlines) {
		$got = $chart.SeriesCollection([int]$want.trendlines.series).Trendlines().Count
		if ($got -ne [int]$want.trendlines.count) { $problems += "$where has $got trendlines, expected $($want.trendlines.count)" }
	}
	if ($want.errorBars) {
		$series = $chart.SeriesCollection([int]$want.errorBars.series)
		if (-not $series.HasErrorBars) { $problems += "$where series $($want.errorBars.series) has no error bars" }
	}
	foreach ($s in @($want.pictureSeries)) {
		if ($null -eq $s) { continue }
		$type = [int]$chart.SeriesCollection([int]$s).Format.Fill.Type
		if ($pictureTypes -notcontains $type) { $problems += "$where series $s fill type is $type, expected a picture" }
	}
	if ($want.pictureArea) {
		$type = [int]$chart.PlotArea.Format.Fill.Type
		if ($pictureTypes -notcontains $type) { $problems += "$where plot area fill type is $type, expected a picture" }
	}
	return $problems
}

function Get-RepairLogs { Get-ChildItem ([IO.Path]::GetTempPath()) -Filter 'error*.xml' -ErrorAction SilentlyContinue | ForEach-Object FullName }

$results = @()
foreach ($property in $manifest.PSObject.Properties) {
	$name = $property.Name
	$want = $property.Value
	$path = Join-Path $OutDir $name
	$app = $null
	$file = $null
	$status = 'ok'
	$detail = ''
	$version = ''
	try {
		$charts = @()
		switch ($want.host) {
			'word' {
				$app = New-Object -ComObject Word.Application
				$app.Visible = $false
				$app.DisplayAlerts = 0
				$names = [string[]]@('FileName', 'ConfirmConversions', 'ReadOnly', 'AddToRecentFiles', 'OpenAndRepair', 'NoEncodingDialog')
				$values = [object[]]@([string]$path, $false, $false, $false, $false, $true)
				$docs = $app.Documents
				$file = [System.__ComObject].InvokeMember('Open', [Reflection.BindingFlags]::InvokeMethod, $null, $docs, $values, $null, $null, $names)
				for ($i = 1; $i -le $file.InlineShapes.Count; $i++) {
					$shape = $file.InlineShapes.Item($i)
					if ($shape.HasChart -eq -1 -or $shape.HasChart -eq $true) { $charts += $shape.Chart }
				}
			}
			'excel' {
				$app = New-Object -ComObject Excel.Application
				$app.Visible = $false
				$app.DisplayAlerts = $false
				$before = @(Get-RepairLogs)
				$file = $app.Workbooks.Open($path, 0, $true)
				$after = @(Get-RepairLogs | Where-Object { $before -notcontains $_ })
				if ($after.Count) { throw "Excel repaired the workbook: $((Get-Content -Raw $after[0]) -replace '\s+', ' ')" }
				if ([string]$app.ActiveWindow.Caption -match 'Repaired') { throw 'Excel repaired the workbook' }
				$sheet = $file.Worksheets.Item(1)
				for ($i = 1; $i -le $sheet.ChartObjects().Count; $i++) { $charts += $sheet.ChartObjects($i).Chart }
			}
			'powerpoint' {
				$app = New-Object -ComObject PowerPoint.Application
				$app.DisplayAlerts = 1 # ppAlertsNone
				$file = $app.Presentations.Open($path, -1, 0, 0) # ReadOnly, not untitled, no window
				# Only the edited chart's slide: the fixture has other charts.
				foreach ($shape in $file.Slides.Item([int]$want.slide).Shapes) { if ($shape.HasChart -eq -1) { $charts += $shape.Chart } }
			}
		}
		$version = [string]$app.Version
		$expected = @($want.charts)
		if ($charts.Count -ne $expected.Count) { throw "$($charts.Count) chart(s), expected $($expected.Count)" }
		$problems = @()
		for ($c = 0; $c -lt $charts.Count; $c++) { $problems += @(Test-Chart $charts[$c] $expected[$c] "chart $($c + 1)") }
		if ($problems.Count) { $status = 'CONTENT'; $detail = $problems -join '; ' }
		else { $detail = "$($charts.Count) chart(s) match" }
	} catch {
		$status = 'FAILED'
		$detail = $_.Exception.Message
	} finally {
		if ($file) { try { if ($want.host -eq 'powerpoint') { $file.Close() } else { $file.Close($false) } } catch { } }
		if ($app) { try { $app.Quit() } catch { }; [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app) }
		[GC]::Collect(); [GC]::WaitForPendingFinalizers()
	}
	$results += [pscustomobject]@{ File = $name; Host = $want.host; Version = $version; Status = $status; Detail = $detail }
}
$results | ConvertTo-Json -Depth 4 | Set-Content (Join-Path $OutDir 'results.json')
$results | Format-Table -AutoSize -Wrap | Out-String -Width 220 | Write-Host
$failed = @($results | Where-Object Status -ne 'ok')
if ($failed.Count) { throw "FAILED: $(($failed | ForEach-Object File) -join ', ')" }
Write-Host "All $($results.Count) files opened natively and match the manifest." -ForegroundColor Green
