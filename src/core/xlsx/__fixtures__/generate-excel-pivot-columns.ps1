# Requires desktop Excel. Generates PivotTables with two or three column fields, with several
# measures, and with nested rows and custom subtotals together.
# Usage: pwsh -File generate-excel-pivot-columns.ps1
$ErrorActionPreference = 'Stop'
# Layout: 0 compact, 1 tabular, 2 outline. Loc: 1 subtotals at the top, 2 at the bottom.
# Extra: more measures as Field:Function (Excel's xl constants: count -4112, average -4106).
$cases = @(
	@{ name = 'cols2-compact'; layout = 0; loc = 1; rows = 'Region'; cols = 'Category,Year' },
	@{ name = 'cols2-tabular'; layout = 1; loc = 2; rows = 'Region'; cols = 'Category,Year' },
	@{ name = 'cols2-outline'; layout = 2; loc = 1; rows = 'Region'; cols = 'Category,Year' },
	@{ name = 'cols3-compact'; layout = 0; loc = 1; rows = 'Product'; cols = 'Category,Region,Year' },
	@{ name = 'cols2-measures-compact'; layout = 0; loc = 1; rows = 'Region'; cols = 'Category,Year'; extra = @(,@('Amount', 'Count of Amount', -4112)) },
	@{ name = 'cols2-measures-tabular'; layout = 1; loc = 2; rows = 'Region'; cols = 'Category,Year'; extra = @(,@('Amount', 'Average of Amount', -4106)) },
	@{ name = 'cols2-rows2-compact-bottom'; layout = 0; loc = 2; rows = 'Region,Product'; cols = 'Category,Year' },
	@{ name = 'cols2-rows2-subtotals'; layout = 1; loc = 2; rows = 'Category,Region'; cols = 'Year,Product'; subs = @('010100000000') }
)
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
Add-Type -AssemblyName System.IO.Compression.FileSystem
try {
	foreach ($case in $cases) {
		$out = Join-Path $PSScriptRoot ('excel-pivot-' + $case.name + '.xlsx')
		$wb = $excel.Workbooks.Add()
		$data = $wb.Worksheets.Item(1)
		$data.Name = 'Data'
		$rows = @(
			@('Category', 'Region', 'Product', 'Year', 'Amount', 'Mixed'),
			@('Fruit', 'East', 'Apple', '2025', 10, 1),
			@('Fruit', 'East', 'Pear', '2026', 7.5, 'x'),
			@('Fruit', 'West', 'Apple', '2025', 20, 4),
			@('Fruit', 'West', 'Plum', '2026', 2.25, 5),
			@('Veg', 'East', 'Carrot', '2025', 5, 'y'),
			@('Veg', 'East', 'Leek', '2026', 9, 7),
			@('Veg', 'West', 'Carrot', '2025', 15, 8),
			@('Veg', 'West', 'Carrot', '2026', 4, 9),
			@('Fruit', 'East', 'Apple', '2026', 3, 2)
		)
		for ($r = 0; $r -lt $rows.Count; $r++) {
			for ($c = 0; $c -lt 6; $c++) {
				$v = $rows[$r][$c]
				if ($v -is [string]) { $data.Cells.Item($r + 1, $c + 1).Value2 = [string]$v } else { $data.Cells.Item($r + 1, $c + 1).Value2 = [double]$v }
			}
		}
		$summary = $wb.Worksheets.Add()
		$summary.Name = 'Summary'
		$cache = $wb.PivotCaches().Create(1, 'Data!R1C1:R10C6', 6)
		$pivot = $cache.CreatePivotTable($summary.Range('A3'), 'ColumnsPivot')
		$names = $case.rows.Split(',')
		foreach ($name in $names) { $pivot.PivotFields($name).Orientation = 1 }
		foreach ($name in $case.cols.Split(',')) { $pivot.PivotFields($name).Orientation = 2 }
		$pivot.RowAxisLayout($case.layout)
		[void]$pivot.AddDataField($pivot.PivotFields('Amount'), 'Data', -4157)
		foreach ($more in $case.extra) { [void]$pivot.AddDataField($pivot.PivotFields($more[0]), $more[1], $more[2]) }
		foreach ($name in $names) { $pivot.PivotFields($name).LayoutSubtotalLocation = $case.loc }
		for ($i = 0; $i -lt $case.subs.Count; $i++) {
			if ($case.subs[$i]) {
				$flags = @(); foreach ($ch in $case.subs[$i].ToCharArray()) { $flags += ($ch -eq '1') }
				$pivot.PivotFields($names[$i]).Subtotals = $flags
			}
		}
		$wb.SaveAs($out, 51)
		$wb.Close($false)
		$zip = [IO.Compression.ZipFile]::Open($out, 'Update')
		try {
			foreach ($part in @('docProps/core.xml', 'xl/workbook.xml', 'xl/pivotCache/pivotCacheDefinition1.xml')) {
				$entry = $zip.GetEntry($part)
				if (-not $entry) { continue }
				$reader = [IO.StreamReader]::new($entry.Open()); $text = $reader.ReadToEnd(); $reader.Dispose()
				$text = [regex]::Replace($text, '<cp:lastModifiedBy>[^<]*</cp:lastModifiedBy>', '<cp:lastModifiedBy>Fixture Author</cp:lastModifiedBy>')
				$text = [regex]::Replace($text, '<dc:creator>[^<]*</dc:creator>', '<dc:creator>Fixture Author</dc:creator>')
				$text = [regex]::Replace($text, '<mc:AlternateContent[^>]*><mc:Choice Requires="x15"><x15ac:absPath[^>]*/></mc:Choice></mc:AlternateContent>', '')
				$text = [regex]::Replace($text, 'refreshedBy="[^"]*"', 'refreshedBy="Fixture Author"')
				$entry.Delete()
				$writer = [IO.StreamWriter]::new($zip.CreateEntry($part).Open(), [Text.UTF8Encoding]::new($false)); $writer.Write($text); $writer.Dispose()
			}
		} finally { $zip.Dispose() }
		Write-Host $out
	}
} finally {
	$excel.Quit()
	[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
}
