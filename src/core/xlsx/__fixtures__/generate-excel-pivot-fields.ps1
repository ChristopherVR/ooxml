# Requires desktop Excel. Generates PivotTables with three row fields in every layout, with the
# subtotals set per field (None, or custom functions), and with a column field.
# Usage: pwsh -File generate-excel-pivot-fields.ps1
$ErrorActionPreference = 'Stop'
# Layout: 0 compact, 1 tabular, 2 outline. Loc: 1 subtotals at the top, 2 at the bottom.
# Subtotals: one 12-flag string per row field (Automatic, Sum, Count, Average, Max, Min, Product,
# Count Numbers, StdDev, StdDevp, Var, Varp).
$cases = @(
	@{ name = 'rows3-compact-top'; layout = 0; loc = 1; rows = 'Category,Region,Product' },
	@{ name = 'rows3-compact-bottom'; layout = 0; loc = 2; rows = 'Category,Region,Product' },
	@{ name = 'rows3-outline-top'; layout = 2; loc = 1; rows = 'Category,Region,Product' },
	@{ name = 'rows3-outline-bottom'; layout = 2; loc = 2; rows = 'Category,Region,Product' },
	@{ name = 'rows3-outline-top-repeat'; layout = 2; loc = 1; rows = 'Category,Region,Product'; repeat = $true },
	@{ name = 'rows3-outline-bottom-repeat'; layout = 2; loc = 2; rows = 'Category,Region,Product'; repeat = $true },
	@{ name = 'rows3-tabular'; layout = 1; loc = 2; rows = 'Category,Region,Product' },
	@{ name = 'rows3-tabular-repeat'; layout = 1; loc = 2; rows = 'Category,Region,Product'; repeat = $true },
	@{ name = 'rows3-tabular-columns'; layout = 1; loc = 2; rows = 'Category,Region,Product'; cols = 'Year' },
	@{ name = 'subtotals-all'; layout = 1; loc = 2; rows = 'Category,Region'; subs = @('011111111111') },
	@{ name = 'subtotals-some-compact'; layout = 0; loc = 1; rows = 'Category,Region,Product'; subs = @('000101000000', '010100000000') },
	@{ name = 'subtotals-none-compact'; layout = 0; loc = 1; rows = 'Category,Region,Product'; subs = @('', '000000000000') },
	@{ name = 'subtotals-count-field'; layout = 1; loc = 2; rows = 'Category,Region'; subs = @('010100000000'); fn = -4112 }
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
		$pivot = $cache.CreatePivotTable($summary.Range('A3'), 'FieldsPivot')
		$names = $case.rows.Split(',')
		foreach ($name in $names) { $pivot.PivotFields($name).Orientation = 1 }
		if ($case.cols) { $pivot.PivotFields($case.cols).Orientation = 2 }
		$pivot.RowAxisLayout($case.layout)
		$function = if ($case.fn) { $case.fn } else { -4157 }
		[void]$pivot.AddDataField($pivot.PivotFields('Amount'), 'Data', $function)
		foreach ($name in $names) { $pivot.PivotFields($name).LayoutSubtotalLocation = $case.loc }
		for ($i = 0; $i -lt $case.subs.Count; $i++) {
			if ($case.subs[$i]) {
				$flags = @(); foreach ($ch in $case.subs[$i].ToCharArray()) { $flags += ($ch -eq '1') }
				$pivot.PivotFields($names[$i]).Subtotals = $flags
			}
		}
		if ($case.repeat) { $pivot.RepeatAllLabels(2) }
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
