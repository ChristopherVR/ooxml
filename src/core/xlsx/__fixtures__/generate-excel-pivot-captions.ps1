# Requires desktop Excel. Generates PivotTables whose captions are in another language.
# Usage: pwsh -File generate-excel-pivot-captions.ps1 [-Dump] [-Only name,name]
#
# Only the English Excel is installed (no Office language pack), so its UI language cannot be
# switched. What Excel lets a user rename is set through COM exactly as a localized Excel stores
# it: the row and column header (rowHeaderCaption, colHeaderCaption), the grand total
# (grandTotalCaption), the Values caption (dataCaption) and the data field names. The words only
# the UI language supplies (the "Total" of a subtotal row, the "Total" in front of a measure's
# grand total and "(All)") are then relabelled in the shared strings, the way a localized Excel
# writes them into the cached cells.
param([switch]$Dump, [string[]]$Only = @())
$ErrorActionPreference = 'Stop'
# sub / measureTotal are the templates of the relabelled words, {0} being the item or measure.
$cases = @(
	@{ name = 'captions-de-compact'; layout = 0; rows = 'Category,Region'; measures = @(,@('Summe von Amount', 'Amount', -4157));
		rowHeader = 'Zeilenbeschriftungen'; grand = 'Gesamtergebnis'; values = 'Werte'; sub = '{0} Ergebnis'; measureTotal = 'Gesamt {0}'; all = '(Alle)' },
	@{ name = 'captions-de-columns'; layout = 0; rows = 'Category'; cols = 'Year'; measures = @(@('Summe von Amount', 'Amount', -4157), @('Anzahl von Amount', 'Amount', -4112));
		rowHeader = 'Zeilenbeschriftungen'; colHeader = 'Spaltenbeschriftungen'; grand = 'Gesamtergebnis'; values = 'Werte'; sub = '{0} Ergebnis'; measureTotal = 'Gesamt {0}'; all = '(Alle)' },
	@{ name = 'captions-es-tabular-page'; layout = 1; rows = 'Category,Region'; cols = 'Year'; page = 'Product'; measures = @(,@('Suma de Amount', 'Amount', -4157));
		rowHeader = 'Etiquetas de fila'; colHeader = 'Etiquetas de columna'; grand = 'Total general'; values = 'Valores'; sub = '{0} Total'; measureTotal = 'Total {0}'; all = '(Todas)' },
	@{ name = 'captions-fr-rows-values'; layout = 1; rows = 'Category,Values,Region'; measures = @(@('Somme de Amount', 'Amount', -4157), @('Nombre de Amount', 'Amount', -4112));
		rowHeader = 'Étiquettes de lignes'; grand = 'Total général'; values = 'Valeurs'; sub = '{0} Total'; measureTotal = 'Total {0}'; all = '(Tous)' },
	@{ name = 'captions-fr-rows-values-compact-page'; layout = 0; rows = 'Category,Region,Values'; page = 'Year'; measures = @(@('Somme de Amount', 'Amount', -4157), @('Nombre de Amount', 'Amount', -4112));
		rowHeader = 'Étiquettes de lignes'; grand = 'Total général'; values = 'Valeurs'; sub = '{0} Total'; measureTotal = 'Total {0}'; all = '(Tous)' },
	# Item display names (PivotItem.Caption, stored as item/@n) with the English words around them.
	@{ name = 'captions-items-tabular'; layout = 1; rows = 'Category,Region'; cols = 'Year'; measures = @(,@('Sum of Amount', 'Amount', -4157));
		rowHeader = 'Row Labels'; grand = 'Grand Total'; values = 'Values'; sub = '{0} Total'; measureTotal = 'Total {0}'; all = '(All)';
		rename = @(@('Category', 'Fruit', 'Obst'), @('Region', 'East', 'Osten')) },
	@{ name = 'captions-items-compact-page'; layout = 0; rows = 'Category,Region'; page = 'Year'; measures = @(,@('Sum of Amount', 'Amount', -4157));
		rowHeader = 'Row Labels'; grand = 'Grand Total'; values = 'Values'; sub = '{0} Total'; measureTotal = 'Total {0}'; all = '(All)';
		rename = @(@('Category', 'Fruit', 'Obst'), @('Region', 'East', 'Osten')) },
	@{ name = 'captions-zh-compact'; layout = 0; rows = 'Category,Region'; cols = 'Year'; measures = @(,@('求和项:Amount', 'Amount', -4157));
		rowHeader = '行标签'; colHeader = '列标签'; grand = '总计'; values = '数值'; sub = '{0} 汇总'; measureTotal = '总计 {0}'; all = '(全部)' }
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
		$pivot = $cache.CreatePivotTable($summary.Range('A3'), 'CaptionsPivot')
		$all = $case.rows.Split(',')
		$names = @($all | Where-Object { $_ -ne 'Values' })
		foreach ($name in $names) { $pivot.PivotFields($name).Orientation = 1 }
		if ($case.cols) { foreach ($name in $case.cols.Split(',')) { $pivot.PivotFields($name).Orientation = 2 } }
		if ($case.page) { $pivot.PivotFields($case.page).Orientation = 3 }
		$pivot.RowAxisLayout($case.layout)
		foreach ($m in $case.measures) { [void]$pivot.AddDataField($pivot.PivotFields($m[1]), $m[0], $m[2]) }
		if ($all -contains 'Values') {
			$pivot.DataPivotField.Orientation = 1
			$pivot.DataPivotField.Position = [array]::IndexOf($all, 'Values') + 1
		}
		foreach ($name in $names) { $pivot.PivotFields($name).LayoutSubtotalLocation = 2 }
		foreach ($r in @($case.rename)) { if ($r) { $pivot.PivotFields($r[0]).PivotItems($r[1]).Caption = $r[2] } }
		$pivot.GrandTotalName = $case.grand
		if ($case.layout -eq 0) {
			$pivot.CompactLayoutRowHeader = $case.rowHeader
			if ($case.colHeader) { $pivot.CompactLayoutColumnHeader = $case.colHeader }
		}
		if ($case.measures.Count -gt 1) { $pivot.DataPivotField.Caption = $case.values }
		if ($Dump) {
			Write-Host "== $($case.name) $($pivot.TableRange2.Address())"
			foreach ($row in $pivot.TableRange2.Rows) {
				$line = @(); foreach ($cell in $row.Cells) { $line += [string]$cell.Value2 }
				Write-Host ($line -join ' | ')
			}
		}
		$wb.SaveAs($out, 51)
		$wb.Close($false)
		$zip = [IO.Compression.ZipFile]::Open($out, 'Update')
		try {
			foreach ($part in @('docProps/core.xml', 'xl/workbook.xml', 'xl/pivotCache/pivotCacheDefinition1.xml', 'xl/sharedStrings.xml')) {
				$entry = $zip.GetEntry($part)
				if (-not $entry) { continue }
				$reader = [IO.StreamReader]::new($entry.Open(), [Text.Encoding]::UTF8); $text = $reader.ReadToEnd(); $reader.Dispose()
				$text = [regex]::Replace($text, '<cp:lastModifiedBy>[^<]*</cp:lastModifiedBy>', '<cp:lastModifiedBy>Fixture Author</cp:lastModifiedBy>')
				$text = [regex]::Replace($text, '<dc:creator>[^<]*</dc:creator>', '<dc:creator>Fixture Author</dc:creator>')
				$text = [regex]::Replace($text, '<mc:AlternateContent[^>]*><mc:Choice Requires="x15"><x15ac:absPath[^>]*/></mc:Choice></mc:AlternateContent>', '')
				$text = [regex]::Replace($text, 'refreshedBy="[^"]*"', 'refreshedBy="Fixture Author"')
				if ($part -eq 'xl/sharedStrings.xml') {
					$sub = $case.sub; $mt = $case.measureTotal; $allText = $case.all
					# The words the UI language supplies: subtotal rows, measure totals and (All).
					$text = [regex]::Replace($text, '(<t[^>]*>)([^<]+) Total(</t>)', [Text.RegularExpressions.MatchEvaluator]{ param($m) $m.Groups[1].Value + $sub.Replace('{0}', $m.Groups[2].Value) + $m.Groups[3].Value })
					$text = [regex]::Replace($text, '(<t[^>]*>)Total ([^<]+)(</t>)', [Text.RegularExpressions.MatchEvaluator]{ param($m) $m.Groups[1].Value + $mt.Replace('{0}', $m.Groups[2].Value) + $m.Groups[3].Value })
					$text = $text.Replace('>(All)<', '>' + $allText + '<')
				}
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
