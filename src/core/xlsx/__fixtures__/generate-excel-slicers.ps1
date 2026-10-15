param([ValidateSet('table','pivot','timeline')][string]$Kind='table', [switch]$Range)
# Requires desktop Excel. Generates slicer and timeline references (Excel 2013+ slicer caches).
$ErrorActionPreference = 'Stop'
$out = Join-Path $PSScriptRoot ('excel-slicer-' + $Kind + $(if ($Range) { '-range' } else { '' }) + '.xlsx')
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$missing = [Type]::Missing
try {
	$wb = $excel.Workbooks.Add()
	$data = $wb.Worksheets.Item(1)
	$data.Name = 'Data'
	$rows = @(@('Category','Region','Amount','Day'),
		@('Fruit','East',10,45658), @('Fruit','West',20,45689), @('Veg','East',5,45717),
		@('Veg','West',15,45748), @('Fruit','East',7,45778), @('Nuts','West',9,45809))
	for ($r = 0; $r -lt $rows.Count; $r++) {
		for ($c = 0; $c -lt 4; $c++) {
			$v = $rows[$r][$c]
			if ($v -is [int]) { $data.Cells.Item($r + 1,$c + 1).Value2 = [double]$v } else { $data.Cells.Item($r + 1,$c + 1).Value2 = [string]$v }
		}
	}
	$data.Range('D2:D7').NumberFormat = 'yyyy-mm-dd'
	$list = $data.ListObjects.Add(1, $data.Range('A1:D7'), $missing, 1)
	$list.Name = 'Sales'
	if ($Kind -eq 'table') {
		$sc = $wb.SlicerCaches.Add2($list, 'Category')
		[void]$sc.Slicers.Add($data, $missing, 'Category', 'Category', 10, 400, 140, 160)
		$sc.SlicerItems.Item('Veg').Selected = $false
		$sc2 = $wb.SlicerCaches.Add2($list, 'Region')
		$s2 = $sc2.Slicers.Add($data, $missing, 'Region', 'Region', 190, 400, 140, 100)
		$s2.NumberOfColumns = 2
	} else {
		$summary = $wb.Worksheets.Add()
		$summary.Name = 'Summary'
		# A range source (not the table name): the pivot reader admits worksheet-range caches only.
		$last = if ($Kind -eq 'pivot') { 'R7C3' } else { 'R7C4' }
		$cache = $wb.PivotCaches().Create(1, ('Data!R1C1:' + $last), 6)
		$pivot = $cache.CreatePivotTable($summary.Range('A3'), 'SalesPivot')
		$pivot.PivotFields('Category').Orientation = 1
		if ($Kind -eq 'pivot') { $pivot.PivotFields('Region').Orientation = 2 }
		[void]$pivot.AddDataField($pivot.PivotFields('Amount'), 'Sum of Amount', -4157)
		if ($Kind -eq 'pivot') {
			$sc = $wb.SlicerCaches.Add2($pivot, 'Category')
			[void]$sc.Slicers.Add($summary, $missing, 'Category', 'Category', 10, 300, 140, 160)
			$sc.SlicerItems.Item('Veg').Selected = $false
		} else {
			$sc = $wb.SlicerCaches.Add2($pivot, 'Day', 'Day', 2)
			if ($Range) { $sc.TimelineState.SetFilterDateRange([datetime]'2025-02-01', [datetime]'2025-04-30') }
			[void]$sc.Slicers.Add($summary, $missing, 'Day', 'Day', 10, 300, 360, 120)
		}
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
