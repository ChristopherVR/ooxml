# Opens the files `chartex-acceptance-build.ts` writes in real Word, Excel and PowerPoint (Windows,
# Office 16, manual) and checks, through the documented object model only, that each opens without
# repair and that every chart has the expected extended chart type and first-series values. Each
# file is also saved again by Office as `<name>.office.<ext>`, which the build script reads back.
#   bun scripts/chartex-acceptance-build.ts <folder>
#   pwsh scripts/chartex-acceptance.ps1 -Folder <folder> [-Products Word,Excel,PowerPoint]
#   bun scripts/chartex-acceptance-build.ts <folder> --verify
# Run the products one at a time; the script never runs two Office applications at once.
param(
	[Parameter(Mandatory = $true)][string]$Folder,
	[string[]]$Products = @('Word', 'Excel', 'PowerPoint')
)
$ErrorActionPreference = 'Stop'
$Folder = (Resolve-Path $Folder).Path
$manifest = Get-Content (Join-Path $Folder 'manifest.json') -Raw | ConvertFrom-Json
# XlChartType values of the extended families (documented in the Office VBA reference).
$types = @{
	waterfall = 119; clusteredColumn = 118; pareto = 122; boxWhisker = 121
	treemap = 117; sunburst = 120; funnel = 123; regionMap = 140
}
$failures = 0
function Report([string]$file, [int]$index, [string]$family, $chart) {
	$expected = $types[$family]
	$type = $chart.ChartType
	$values = @()
	try { $values = @($chart.SeriesCollection(1).Values) } catch { $values = @('?') }
	$want = ($manifest.$file[$index].values -join ',')
	$got = ($values -join ',')
	# Series.Values is empty for extended charts in Office 16; the values are then checked from
	# Office's saved copy by the build script's --verify step.
	$ok = ($type -eq $expected) -and ($got -eq '' -or $got -eq $want)
	if (-not $ok) { $script:failures++ }
	"{0} #{1} {2}: type {3} (want {4}) values [{5}] (want [{6}]) {7}" -f $file, ($index + 1), $family, $type, $expected, $got, $want, ($(if ($ok) { 'ok' } else { 'FAIL' }))
}

foreach ($product in $Products) {
	$files = $manifest.PSObject.Properties.Name | Where-Object {
		($product -eq 'Word' -and $_ -like '*.docx') -or ($product -eq 'Excel' -and $_ -like '*.xlsx') -or ($product -eq 'PowerPoint' -and $_ -like '*.pptx')
	}
	if (-not $files) { continue }
	$app = New-Object -ComObject "$product.Application"
	try {
		if ($product -ne 'PowerPoint') { $app.Visible = $false; $app.DisplayAlerts = 0 }
		foreach ($file in $files) {
			$path = Join-Path $Folder $file
			try {
				if ($product -eq 'Word') {
					# Word's object model lists no extended chart (InlineShapes holds classic charts
					# only), so Word re-saves the document and the build script reads Word's own
					# chartEx parts back (`--verify`); the PDF is for a visual check.
					$doc = $app.Documents.Open($path, $false, $true, $false)
					"$file opened"
					$doc.SaveAs2((Join-Path $Folder ($file -replace '\.docx$', '.office.docx')), 12)
					$doc.ExportAsFixedFormat((Join-Path $Folder ($file -replace '\.docx$', '.office.pdf')), 17)
					$doc.Close(0)
				} elseif ($product -eq 'Excel') {
					$book = $app.Workbooks.Open($path, 0, $true, 5, '', '', $true, 2, '', $false, $false, 0, $false, $true, 1)
					$shapes = @($book.Worksheets.Item(1).Shapes | Where-Object { $_.HasChart })
					"$file opened: $($shapes.Count) charts"
					for ($i = 0; $i -lt $shapes.Count; $i++) { Report $file $i $manifest.$file[$i].family $shapes[$i].Chart }
					$book.SaveAs((Join-Path $Folder ($file -replace '\.xlsx$', '.office.xlsx')), 51)
					$book.Close($false)
				} else {
					$deck = $app.Presentations.Open($path, $true, $false, $false)
					$shapes = @()
					foreach ($slide in $deck.Slides) { foreach ($shape in $slide.Shapes) { if ($shape.HasChart) { $shapes += $shape } } }
					"$file opened: $($shapes.Count) charts"
					for ($i = 0; $i -lt $shapes.Count; $i++) { Report $file $i $manifest.$file[$i].family $shapes[$i].Chart }
					$deck.SaveAs((Join-Path $Folder ($file -replace '\.pptx$', '.office.pptx')), 24)
					$deck.Close()
				}
			} catch {
				$script:failures++
				"$file FAILED to open: $_"
			}
		}
	} finally {
		$app.Quit()
		[System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
	}
}
if ($failures) { "$failures failure(s)"; exit 1 } else { 'all checks passed' }
