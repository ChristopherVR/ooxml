# Requires desktop Excel. Generates PivotTables whose row field is a date field grouped by
# seconds, minutes, hours, days, months, quarters and years (Group Selection).
# Usage: pwsh -File generate-excel-pivot-dates.ps1
$ErrorActionPreference = 'Stop'
# Periods: one flag per unit in the order Seconds, Minutes, Hours, Days, Months, Quarters, Years.
# By: the number of days (only when grouping by days alone). Start and End: serials, else automatic.
# Layout: 0 compact, 1 tabular, 2 outline. Times: the dates carry a time of day.
$cases = @(
	@{ name = 'months-years'; periods = '0000101'; layout = 0 },
	@{ name = 'months-quarters-years-tabular'; periods = '0000111'; layout = 1 },
	@{ name = 'quarters-years'; periods = '0000011'; layout = 0 },
	@{ name = 'years'; periods = '0000001'; layout = 0 },
	@{ name = 'days-7'; periods = '0001000'; layout = 0; by = 7 },
	@{ name = 'days-1'; periods = '0001000'; layout = 0; by = 1 },
	@{ name = 'time'; periods = '1110000'; layout = 0; times = $true },
	@{ name = 'months-custom'; periods = '0000100'; layout = 2; start = 45300; end = 45500 },
	@{ name = 'year-end'; periods = '0000101'; layout = 0; dates = '2024-03-10,2024-05-01,2024-12-31,2024-06-06,2024-01-02,2024-08-15,2024-09-30,2024-10-31' }
)
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
Add-Type -AssemblyName System.IO.Compression.FileSystem
try {
	foreach ($case in $cases) {
		$out = Join-Path $PSScriptRoot ('excel-pivot-group-date-' + $case.name + '.xlsx')
		$wb = $excel.Workbooks.Add()
		$data = $wb.Worksheets.Item(1)
		$data.Name = 'Data'
		$data.Cells.Item(1, 1).Value2 = [string]'Date'
		$data.Cells.Item(1, 2).Value2 = [string]'Region'
		$data.Cells.Item(1, 3).Value2 = [string]'Amount'
		$dates = if ($case.dates) { $case.dates.Split(',') } else { @('2023-11-15', '2023-12-20', '2024-01-05', '2024-02-14', '2024-02-28', '2024-04-09', '2024-07-30', '2024-11-02') }
		$regions = @('East', 'West', 'East', 'West', 'East', 'West', 'East', 'West')
		$amounts = @(10, 20, 5, 15, 7, 9, 12, 3)
		for ($i = 0; $i -lt 8; $i++) {
			$r = $i + 2
			$time = if ($case.times) { $i * 0.0123456 } else { 0 }
			$data.Cells.Item($r, 1).Value2 = ([datetime]$dates[$i]).ToOADate() + $time
			$data.Cells.Item($r, 1).NumberFormat = $(if ($case.times) { 'm/d/yyyy h:mm:ss' } else { 'm/d/yyyy' })
			$data.Cells.Item($r, 2).Value2 = [string]$regions[$i]
			$data.Cells.Item($r, 3).Value2 = [double]$amounts[$i]
		}
		$summary = $wb.Worksheets.Add()
		$summary.Name = 'Summary'
		$cache = $wb.PivotCaches().Create(1, 'Data!R1C1:R9C3', 6)
		$pivot = $cache.CreatePivotTable($summary.Range('A3'), 'DatePivot')
		$pivot.PivotFields('Date').Orientation = 1
		[void]$pivot.AddDataField($pivot.PivotFields('Amount'), 'Sum of Amount', -4157)
		$periods = @(); foreach ($ch in $case.periods.ToCharArray()) { $periods += ($ch -eq '1') }
		$start = $true; if ($case.start) { $start = [double]$case.start }
		$end = $true; if ($case.end) { $end = [double]$case.end }
		$by = [Type]::Missing; if ($case.by) { $by = [int]$case.by }
		[void]$pivot.PivotFields('Date').DataRange.Cells(1).Group($start, $end, $by, $periods)
		$pivot.RowAxisLayout($case.layout)
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
