# Requires desktop Excel. Generates PivotTables with Values on rows and any number of row and
# column fields: Values at every position among three row fields, nested columns, and custom
# subtotals on row and column fields.
# Usage: pwsh -File generate-excel-pivot-row-values.ps1 [-Dump] [-Only name,name]
param([switch]$Dump, [string[]]$Only = @())
$ErrorActionPreference = 'Stop'
# Layout: 0 compact, 1 tabular, 2 outline. Loc: 1 subtotals at the top, 2 at the bottom.
# rows names the row fields with Values at its place ('Category,Values,Region'). Subtotals: one
# 12-flag string per field (Automatic, Sum, Count, Average, Max, Min, Product, Count Numbers,
# StdDev, StdDevp, Var, Varp); rowSubs and colSubs apply to the fields in order.
$cases = @(
	@{ name = 'rv3-last-compact-top'; layout = 0; loc = 1; rows = 'Category,Region,Product,Values' },
	@{ name = 'rv3-first-compact-top'; layout = 0; loc = 1; rows = 'Values,Category,Region,Product' },
	@{ name = 'rv3-first-compact-bottom'; layout = 0; loc = 2; rows = 'Values,Category,Region,Product' },
	@{ name = 'rv3-second-compact-top'; layout = 0; loc = 1; rows = 'Category,Values,Region,Product' },
	@{ name = 'rv3-third-compact-top'; layout = 0; loc = 1; rows = 'Category,Region,Values,Product' },
	@{ name = 'rv3-last-compact-bottom'; layout = 0; loc = 2; rows = 'Category,Region,Product,Values' },
	@{ name = 'rv3-last-outline-top'; layout = 2; loc = 1; rows = 'Category,Region,Product,Values' },
	@{ name = 'rv3-last-outline-bottom'; layout = 2; loc = 2; rows = 'Category,Region,Product,Values' },
	@{ name = 'rv3-last-tabular'; layout = 1; loc = 2; rows = 'Category,Region,Product,Values' },
	@{ name = 'rv3-first-outline-top'; layout = 2; loc = 1; rows = 'Values,Category,Region,Product' },
	@{ name = 'rv3-first-outline-bottom'; layout = 2; loc = 2; rows = 'Values,Category,Region,Product' },
	@{ name = 'rv3-first-tabular'; layout = 1; loc = 2; rows = 'Values,Category,Region,Product' },
	@{ name = 'rv3-second-outline-bottom'; layout = 2; loc = 2; rows = 'Category,Values,Region,Product' },
	@{ name = 'rv3-second-tabular'; layout = 1; loc = 2; rows = 'Category,Values,Region,Product' },
	@{ name = 'rv3-third-outline-top'; layout = 2; loc = 1; rows = 'Category,Region,Values,Product' },
	@{ name = 'rv3-third-tabular'; layout = 1; loc = 2; rows = 'Category,Region,Values,Product' },
	@{ name = 'rv2-cols2-last-compact'; layout = 0; loc = 1; rows = 'Category,Region,Values'; cols = 'Year,Product' },
	@{ name = 'rv2-cols2-first-tabular'; layout = 1; loc = 2; rows = 'Values,Category,Region'; cols = 'Year,Product' },
	@{ name = 'rv1-cols2-compact'; layout = 0; loc = 1; rows = 'Region,Values'; cols = 'Category,Year' },
	@{ name = 'rv1-cols2-tabular'; layout = 1; loc = 2; rows = 'Values,Region'; cols = 'Category,Year' },
	@{ name = 'rv1-first-compact'; layout = 0; loc = 1; rows = 'Values,Region' },
	@{ name = 'rv1-first-outline-cols'; layout = 2; loc = 1; rows = 'Values,Region'; cols = 'Category' },
	@{ name = 'rv3-last-outline-top-repeat'; layout = 2; loc = 1; rows = 'Category,Region,Product,Values'; repeat = $true },
	@{ name = 'rv3-first-tabular-repeat'; layout = 1; loc = 2; rows = 'Values,Category,Region,Product'; repeat = $true },
	@{ name = 'rv3-second-outline-top-repeat'; layout = 2; loc = 1; rows = 'Category,Values,Region,Product'; repeat = $true },
	@{ name = 'rv3-third-tabular-repeat'; layout = 1; loc = 2; rows = 'Category,Region,Values,Product'; repeat = $true },
	@{ name = 'rv2-cols2-second-outline'; layout = 2; loc = 1; rows = 'Category,Values,Region'; cols = 'Year,Product' },
	@{ name = 'rvs-first-tabular-cols'; layout = 1; loc = 2; rows = 'Values,Category,Region'; cols = 'Year,Product'; rowSubs = @('011000000000'); colSubs = @('010000000000') },
	@{ name = 'rvs-third-outline'; layout = 2; loc = 1; rows = 'Category,Region,Values,Product'; rowSubs = @('010000000000', '000100000000') },
	@{ name = 'rvs-none-last'; layout = 0; loc = 1; rows = 'Category,Region,Values'; rowSubs = @('000000000000') },
	@{ name = 'rvs-none-first'; layout = 1; loc = 2; rows = 'Values,Category,Region'; rowSubs = @('000000000000') },
	@{ name = 'rvs-last-compact';layout = 0; loc = 1; rows = 'Category,Region,Product,Values'; rowSubs = @('010100000000', '000000000000') },
	@{ name = 'rvs-first-compact-top'; layout = 0; loc = 1; rows = 'Values,Category,Region,Product'; rowSubs = @('010100000000', '000000000000') },
	@{ name = 'rvs-first-compact-bottom'; layout = 0; loc = 2; rows = 'Values,Category,Region,Product'; rowSubs = @('010100000000', '000000000000') },
	@{ name = 'rvs-second-tabular'; layout = 1; loc = 2; rows = 'Category,Values,Region,Product'; rowSubs = @('011000000000', '000000000000') },
	@{ name = 'rvs-cols-compact'; layout = 0; loc = 1; rows = 'Region,Values'; cols = 'Category,Year'; colSubs = @('010100000000') },
	@{ name = 'rvs-rows-cols-tabular'; layout = 1; loc = 2; rows = 'Category,Region,Values'; cols = 'Year,Product'; rowSubs = @('010100000000'); colSubs = @('011000000000') }
)
if ($Only.Count) { $Only = @($Only | ForEach-Object { $_.Split(',') }); $cases = @($cases | Where-Object { $Only -contains $_.name }) }
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
		$pivot = $cache.CreatePivotTable($summary.Range('A3'), 'RowValuesPivot')
		$all = $case.rows.Split(',')
		$names = @($all | Where-Object { $_ -ne 'Values' })
		$cols = if ($case.cols) { $case.cols.Split(',') } else { @() }
		foreach ($name in $names) { $pivot.PivotFields($name).Orientation = 1 }
		foreach ($name in $cols) { $pivot.PivotFields($name).Orientation = 2 }
		$pivot.RowAxisLayout($case.layout)
		[void]$pivot.AddDataField($pivot.PivotFields('Amount'), 'Data', -4157)
		[void]$pivot.AddDataField($pivot.PivotFields('Amount'), 'Count of Amount', -4112)
		$pivot.DataPivotField.Orientation = 1
		$pivot.DataPivotField.Position = [array]::IndexOf($all, 'Values') + 1
		foreach ($name in $names) { $pivot.PivotFields($name).LayoutSubtotalLocation = $case.loc }
		if ($case.repeat) { $pivot.RepeatAllLabels(2) }
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
