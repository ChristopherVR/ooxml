# Requires desktop Excel. Generates PivotTables with date groups (months and years on the rows)
# together with a report filter, hidden items, a slicer and a timeline.
# Usage: pwsh -File generate-excel-pivot-date-filters.ps1 [-Dump] [-Only name,name]
param([switch]$Dump, [string[]]$Only = @())
$ErrorActionPreference = 'Stop'
# Layout: 0 compact, 1 tabular, 2 outline. Periods: Seconds, Minutes, Hours, Days, Months, Quarters, Years.
# page: a report filter field and its item; hide: items hidden on a column field; slicer: a slicer
# on a field with one item deselected; timeline: an inclusive date range on the Date field.
$cases = @(
	@{ name = 'group-page-all'; periods = '0000101'; layout = 0; page = 'Region'; item = '(All)' },
	@{ name = 'group-page-west'; periods = '0000101'; layout = 0; page = 'Region'; item = 'West' },
	@{ name = 'group-page-tabular-quarters'; periods = '0000111'; layout = 1; page = 'Category'; item = 'Fruit' },
	@{ name = 'group-page-columns'; periods = '0000101'; layout = 2; page = 'Category'; item = 'Veg'; cols = 'Region' },
	@{ name = 'group-hidden-items'; periods = '0000101'; layout = 0; cols = 'Region'; hideField = 'Region'; hideItem = 'West' },
	@{ name = 'group-slicer'; periods = '0000101'; layout = 0; slicer = 'Category'; slicerItem = 'Veg' },
	@{ name = 'group-timeline'; periods = '0000101'; layout = 0; timeline = $true; from = '2024-01-01'; to = '2024-06-30' },
	@{ name = 'group-timeline-days'; periods = '0001000'; layout = 0; timeline = $true; from = '2023-12-01'; to = '2024-03-01'; by = 30 },
	@{ name = 'group-page-timeline'; periods = '0000101'; layout = 0; page = 'Region'; item = 'East'; timeline = $true; from = '2024-01-01'; to = '2024-12-31' }
)
if ($Only.Count) { $Only = @($Only | ForEach-Object { $_.Split(',') }); $cases = @($cases | Where-Object { $Only -contains $_.name }) }
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$missing = [Type]::Missing
Add-Type -AssemblyName System.IO.Compression.FileSystem
try {
	foreach ($case in $cases) {
		$out = Join-Path $PSScriptRoot ('excel-pivot-' + $case.name + '.xlsx')
		$wb = $excel.Workbooks.Add()
		$data = $wb.Worksheets.Item(1)
		$data.Name = 'Data'
		foreach ($pair in @(@(1, 'Date'), @(2, 'Region'), @(3, 'Category'), @(4, 'Amount'))) { $data.Cells.Item(1, $pair[0]).Value2 = [string]$pair[1] }
		$dates = @('2023-11-15', '2023-12-20', '2024-01-05', '2024-02-14', '2024-02-28', '2024-04-09', '2024-07-30', '2024-11-02')
		$regions = @('East', 'West', 'East', 'West', 'East', 'West', 'East', 'West')
		$categories = @('Fruit', 'Fruit', 'Veg', 'Veg', 'Fruit', 'Veg', 'Fruit', 'Veg')
		$amounts = @(10, 20, 5, 15, 7, 9, 12, 3)
		for ($i = 0; $i -lt 8; $i++) {
			$r = $i + 2
			$data.Cells.Item($r, 1).Value2 = ([datetime]$dates[$i]).ToOADate()
			$data.Cells.Item($r, 1).NumberFormat = 'm/d/yyyy'
			$data.Cells.Item($r, 2).Value2 = [string]$regions[$i]
			$data.Cells.Item($r, 3).Value2 = [string]$categories[$i]
			$data.Cells.Item($r, 4).Value2 = [double]$amounts[$i]
		}
		$summary = $wb.Worksheets.Add()
		$summary.Name = 'Summary'
		$cache = $wb.PivotCaches().Create(1, 'Data!R1C1:R9C4', 6)
		$pivot = $cache.CreatePivotTable($summary.Range('A3'), 'DateFilterPivot')
		$pivot.PivotFields('Date').Orientation = 1
		if ($case.cols) { $pivot.PivotFields($case.cols).Orientation = 2 }
		[void]$pivot.AddDataField($pivot.PivotFields('Amount'), 'Sum of Amount', -4157)
		$periods = @(); foreach ($ch in $case.periods.ToCharArray()) { $periods += ($ch -eq '1') }
		$by = $missing; if ($case.by) { $by = [int]$case.by }
		[void]$pivot.PivotFields('Date').DataRange.Cells(1).Group($true, $true, $by, $periods)
		$pivot.RowAxisLayout($case.layout)
		if ($case.page) {
			$pivot.PivotFields($case.page).Orientation = 3
			if ($case.item -ne '(All)') { $pivot.PivotFields($case.page).CurrentPage = $case.item }
		}
		if ($case.hideField) { $pivot.PivotFields($case.hideField).PivotItems($case.hideItem).Visible = $false }
		if ($case.slicer) {
			$sc = $wb.SlicerCaches.Add2($pivot, $case.slicer)
			[void]$sc.Slicers.Add($summary, $missing, $case.slicer, $case.slicer, 10, 400, 140, 120)
			$sc.SlicerItems.Item($case.slicerItem).Selected = $false
		}
		if ($case.timeline) {
			$tc = $wb.SlicerCaches.Add2($pivot, 'Date', 'Date', 2)
			$tc.TimelineState.SetFilterDateRange([datetime]$case.from, [datetime]$case.to)
			[void]$tc.Slicers.Add($summary, $missing, 'Date', 'Date', 10, 400, 360, 120)
		}
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
