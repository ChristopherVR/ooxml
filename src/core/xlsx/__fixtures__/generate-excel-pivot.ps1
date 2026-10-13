param([switch]$Hierarchy, [switch]$NoSubtotals, [switch]$RepeatLabels, [switch]$MultiValues, [switch]$ColumnAxis, [ValidateSet('tabular','compact','outline')][string]$Layout='tabular', [ValidateSet('top','bottom')][string]$SubtotalPosition='bottom')
# Requires desktop Excel. Generates a tabular PivotTable reference.
$ErrorActionPreference = 'Stop'
$out = Join-Path $PSScriptRoot 'excel-pivot.xlsx'
if ($Hierarchy) {
	$suffix = if ($NoSubtotals) { '-no-subtotals' } elseif ($RepeatLabels) { '-repeat' } else { '' }
	$out = Join-Path $PSScriptRoot ('excel-pivot-rows' + $suffix + '.xlsx')
	if ($Layout -ne 'tabular') { $out = Join-Path $PSScriptRoot ('excel-pivot-' + $Layout + '-' + $SubtotalPosition + $suffix + '.xlsx') }
}
$excel = New-Object -ComObject Excel.Application
if ($MultiValues) { $out = Join-Path $PSScriptRoot ('excel-pivot-values' + $(if ($Hierarchy) {'-rows'} else {''}) + $(if ($ColumnAxis) {'-columns'} else {''}) + '.xlsx') }
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
	$wb = $excel.Workbooks.Add()
	$data = $wb.Worksheets.Item(1)
	$data.Name = 'Data'
	$rows = @(@('Category','Region','Amount'), @('Fruit','East',10), @('Fruit','West',20), @('Veg','East',5), @('Veg','West',15), @('Fruit','East',7))
	if ($MultiValues) { $rows = @(@('Category','Region','Amount','Quantity','Year'), @('Fruit','East',10,2,2025), @('Fruit','West',20,3,2026), @('Veg','East',5,4,2025), @('Veg','West',15,5,2026), @('Fruit','East',7,6,2026)) }
	for ($r = 0; $r -lt $rows.Count; $r++) {
		for ($c = 0; $c -lt $rows[0].Count; $c++) {
			$v = $rows[$r][$c]
			if ($v -is [int]) { $data.Cells.Item($r + 1,$c + 1).Value2 = [double]$v }
			else { $data.Cells.Item($r + 1,$c + 1).Value2 = [string]$v }
		}
	}
	$summary = $wb.Worksheets.Add()
	$summary.Name = 'Summary'
	$cache = $wb.PivotCaches().Create(1, ('Data!R1C1:R6C'+$rows[0].Count), 6)
	$pivot = $cache.CreatePivotTable($summary.Range('A1'), 'SalesPivot')
	$pivot.PivotFields('Category').Orientation = 1
	if ($Hierarchy) { $pivot.PivotFields('Region').Orientation = 1 } elseif (-not $MultiValues) { $pivot.PivotFields('Region').Orientation = 2 }
	if ($ColumnAxis) { $pivot.PivotFields('Year').Orientation = 2 }
	$pivot.RowAxisLayout($(switch ($Layout) { 'compact' { 0 } 'outline' { 2 } default { 1 } }))
	[void]$pivot.AddDataField($pivot.PivotFields('Amount'), 'Sum of Amount', -4157)
	if ($MultiValues) {
		[void]$pivot.AddDataField($pivot.PivotFields('Quantity'), 'Count of Quantity', -4112)
		[void]$pivot.AddDataField($pivot.PivotFields('Amount'), 'Average of Amount', -4106)
	}
	if ($Hierarchy) {
		$pivot.PivotFields('Category').LayoutSubtotalLocation = if ($SubtotalPosition -eq 'top') { 1 } else { 2 }
		if ($NoSubtotals) {
			$pivot.PivotFields('Category').Subtotals = @($false,$false,$false,$false,$false,$false,$false,$false,$false,$false,$false,$false)
		}
		if ($RepeatLabels) { $pivot.RepeatAllLabels(2) }
	}
	$wb.SaveAs($out, 51)
	$wb.Close($false)
} finally {
	$excel.Quit()
	[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
}

Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [IO.Compression.ZipFile]::Open($out, 'Update')
try {
	foreach ($name in @('docProps/core.xml','xl/workbook.xml','xl/pivotCache/pivotCacheDefinition1.xml')) {
		$entry = $zip.GetEntry($name)
		if (-not $entry) { continue }
		$reader = [IO.StreamReader]::new($entry.Open()); $text = $reader.ReadToEnd(); $reader.Dispose()
		$text = [regex]::Replace($text, '<cp:lastModifiedBy>[^<]*</cp:lastModifiedBy>', '<cp:lastModifiedBy>Fixture Author</cp:lastModifiedBy>')
		$text = [regex]::Replace($text, '<dc:creator>[^<]*</dc:creator>', '<dc:creator>Fixture Author</dc:creator>')
		$text = [regex]::Replace($text, '<mc:AlternateContent[^>]*><mc:Choice Requires="x15"><x15ac:absPath[^>]*/></mc:Choice></mc:AlternateContent>', '')
		$text = [regex]::Replace($text, 'refreshedBy="[^"]*"', 'refreshedBy="Fixture Author"')
		$entry.Delete()
		$writer = [IO.StreamWriter]::new($zip.CreateEntry($name).Open(), [Text.UTF8Encoding]::new($false)); $writer.Write($text); $writer.Dispose()
	}
} finally { $zip.Dispose() }
Get-Item $out | Select-Object Name, Length
