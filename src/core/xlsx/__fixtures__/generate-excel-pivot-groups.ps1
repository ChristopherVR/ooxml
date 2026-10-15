# Requires desktop Excel. Generates a PivotTable whose row field is grouped by number ranges.
$ErrorActionPreference = 'Stop'
$out = Join-Path $PSScriptRoot 'excel-pivot-group-number.xlsx'
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
	$wb = $excel.Workbooks.Add()
	$data = $wb.Worksheets.Item(1)
	$data.Name = 'Data'
	$rows = @(@('Category','Score','Amount'), @('A',3,10), @('B',12,20), @('A',17,5), @('B',25,15), @('A',29,7), @('B',44,9))
	for ($r = 0; $r -lt $rows.Count; $r++) {
		for ($c = 0; $c -lt 3; $c++) {
			$v = $rows[$r][$c]
			if ($v -is [int]) { $data.Cells.Item($r + 1,$c + 1).Value2 = [double]$v } else { $data.Cells.Item($r + 1,$c + 1).Value2 = [string]$v }
		}
	}
	$summary = $wb.Worksheets.Add()
	$summary.Name = 'Summary'
	$cache = $wb.PivotCaches().Create(1, 'Data!R1C1:R7C3', 6)
	$pivot = $cache.CreatePivotTable($summary.Range('A3'), 'GroupPivot')
	$pivot.PivotFields('Score').Orientation = 1
	[void]$pivot.AddDataField($pivot.PivotFields('Amount'), 'Sum of Amount', -4157)
	# Group Score: start 0, end 50, by 10 (the groups are <0, 0-9, 10-19, 20-29, 30-39, 40-50, >50).
	$pivot.PivotFields('Score').DataRange.Cells(1).Group(0, 50, 10)
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
