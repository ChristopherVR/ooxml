# Requires desktop Excel. Generates PivotTables whose column fields have subtotals set to None
# or to custom functions, alone, with several measures and together with row subtotals.
# Usage: pwsh -File generate-excel-pivot-column-subtotals.ps1 [-Dump]
param([switch]$Dump)
$ErrorActionPreference = 'Stop'
# Layout: 0 compact, 1 tabular, 2 outline. Loc: 1 subtotals at the top, 2 at the bottom.
# Subtotals: one 12-flag string per field (Automatic, Sum, Count, Average, Max, Min, Product,
# Count Numbers, StdDev, StdDevp, Var, Varp); colSubs applies to the column fields in order.
$cases = @(
	@{ name = 'colsubs-none-compact'; layout = 0; loc = 1; rows = 'Region'; cols = 'Category,Year'; colSubs = @('000000000000') },
	@{ name = 'colsubs-sum-average-compact'; layout = 0; loc = 1; rows = 'Region'; cols = 'Category,Year'; colSubs = @('010100000000') },
	@{ name = 'colsubs-sum-average-tabular'; layout = 1; loc = 2; rows = 'Region'; cols = 'Category,Year'; colSubs = @('010100000000') },
	@{ name = 'colsubs-count-outline'; layout = 2; loc = 1; rows = 'Region'; cols = 'Category,Year'; colSubs = @('001000000000') },
	@{ name = 'colsubs-three-levels'; layout = 0; loc = 1; rows = 'Product'; cols = 'Category,Region,Year'; colSubs = @('010100000000', '000000000000') },
	@{ name = 'colsubs-measures-compact'; layout = 0; loc = 1; rows = 'Region'; cols = 'Category,Year'; colSubs = @('010100000000'); extra = @(,@('Amount', 'Count of Amount', -4112)) },
	@{ name = 'colsubs-rows-subtotals'; layout = 1; loc = 2; rows = 'Category,Region'; cols = 'Year,Product'; rowSubs = @('010100000000'); colSubs = @('000000000000') },
	@{ name = 'colsubs-rows-custom'; layout = 1; loc = 2; rows = 'Category,Region'; cols = 'Year,Product'; rowSubs = @('010100000000'); colSubs = @('011000000000') },
	@{ name = 'colsubs-rows-custom-compact'; layout = 0; loc = 1; rows = 'Category,Region'; cols = 'Year,Product'; rowSubs = @('010100000000'); colSubs = @('011000000000') },
	@{ name = 'colsubs-rows-same'; layout = 1; loc = 2; rows = 'Category,Region'; cols = 'Year,Product'; rowSubs = @('011000000000'); colSubs = @('011000000000') },
	@{ name = 'colsubs-rows-total'; layout = 1; loc = 2; rows = 'Category,Region'; cols = 'Year,Product'; colSubs = @('011000000000') },
	@{ name = 'colsubs-rows-auto';layout = 0; loc = 1; rows = 'Category,Region'; cols = 'Year,Product'; rowSubs = @('010100000000') },
	@{ name = 'colsubs-measures-other'; layout = 1; loc = 2; rows = 'Region'; cols = 'Category,Year'; colSubs = @('011000000000'); extra = @(,@('Mixed', 'Count of Mixed', -4112)) }
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
		$cols = $case.cols.Split(',')
		foreach ($name in $names) { $pivot.PivotFields($name).Orientation = 1 }
		foreach ($name in $cols) { $pivot.PivotFields($name).Orientation = 2 }
		$pivot.RowAxisLayout($case.layout)
		[void]$pivot.AddDataField($pivot.PivotFields('Amount'), 'Data', -4157)
		foreach ($more in $case.extra) { [void]$pivot.AddDataField($pivot.PivotFields($more[0]), $more[1], $more[2]) }
		foreach ($name in $names) { $pivot.PivotFields($name).LayoutSubtotalLocation = $case.loc }
		foreach ($set in @(@($names, $case.rowSubs), @($cols, $case.colSubs))) {
			for ($i = 0; $i -lt $set[1].Count; $i++) {
				$flags = @(); foreach ($ch in $set[1][$i].ToCharArray()) { $flags += ($ch -eq '1') }
				$pivot.PivotFields($set[0][$i]).Subtotals = $flags
			}
		}
		if ($Dump) {
			Write-Host "== $($case.name) $($pivot.TableRange1.Address())"
			foreach ($row in $pivot.TableRange1.Rows) {
				$line = @(); foreach ($cell in $row.Cells) { $line += [string]$cell.Value2 }
				Write-Host ($line -join ' | ')
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
