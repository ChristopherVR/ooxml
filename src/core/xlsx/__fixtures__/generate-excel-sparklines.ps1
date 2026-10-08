# Generates excel-sparklines.xlsx (sparkline groups authored by Excel) for the sparkline tests.
# Requires desktop Excel (COM automation). Run from any folder:
#   pwsh -File src/core/xlsx/__fixtures__/generate-excel-sparklines.ps1
# The account name is never changed; the author fields Excel stamps are replaced afterwards.
$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$out = Join-Path $here 'excel-sparklines.xlsx'

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
	$wb = $excel.Workbooks.Add()
	$ws = $wb.Worksheets.Item(1)
	$ws.Range('A1:E1').Value2 = [object[]](1, 3, -2, 5, 4)
	$ws.Range('A2').Value2 = 2
	$ws.Range('B2').Value2 = 4
	$ws.Range('D2').Value2 = 1
	$ws.Range('E2').Value2 = 3
	$ws.Range('A3:E3').Value2 = [object[]](1, -1, 1, 1, -1)
	$ws.Range('A4:E4').Value2 = [object[]](2, 8, 6, 3, 7)
	# 1 xlSparkLine: markers, high, low and negative points, a 1.5 pt line.
	$line = $ws.Range('F1').SparklineGroups.Add(1, 'Sheet1!A1:E1')
	$line.LineWeight = 1.5
	$line.Points.Markers.Visible = $true
	$line.Points.Highpoint.Visible = $true
	$line.Points.Lowpoint.Visible = $true
	$line.Points.Negative.Visible = $true
	# 2 xlSparkColumn over a row with an empty cell (shown as a gap).
	$column = $ws.Range('F2').SparklineGroups.Add(2, 'Sheet1!A2:E2')
	$column.Points.Firstpoint.Visible = $true
	$column.Points.Lastpoint.Visible = $true
	# 3 xlSparkColumnStacked100 (win/loss) with negative points.
	$stacked = $ws.Range('F3').SparklineGroups.Add(3, 'Sheet1!A3:E3')
	$stacked.Points.Negative.Visible = $true
	# A line with a custom 0-10 vertical axis and the horizontal axis shown.
	$manual = $ws.Range('F4').SparklineGroups.Add(1, 'Sheet1!A4:E4')
	$manual.Axes.Vertical.MinScaleType = 3
	$manual.Axes.Vertical.CustomMinScaleValue = 0
	$manual.Axes.Vertical.MaxScaleType = 3
	$manual.Axes.Vertical.CustomMaxScaleValue = 10
	$manual.Axes.Horizontal.Axis.Visible = $true
	if (Test-Path $out) { Remove-Item $out -Force }
	$wb.SaveAs($out, 51)
	$wb.Close($false)
} finally {
	$excel.Quit()
	[void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}

# Replace the author fields and the save folder Excel writes into the package.
Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
$zip = [IO.Compression.ZipFile]::Open($out, 'Update')
try {
	$edits = @(
		@{ Part = 'docProps/core.xml'; Find = '<cp:lastModifiedBy>[^<]*</cp:lastModifiedBy>'; With = '<cp:lastModifiedBy>Fixture Author</cp:lastModifiedBy>' }
		@{ Part = 'docProps/core.xml'; Find = '<dc:creator>[^<]*</dc:creator>'; With = '<dc:creator>Fixture Author</dc:creator>' }
		@{ Part = 'xl/workbook.xml'; Find = '<mc:AlternateContent[^>]*><mc:Choice Requires="x15"><x15ac:absPath[^>]*/></mc:Choice></mc:AlternateContent>'; With = '' }
	)
	foreach ($name in ($edits | ForEach-Object { $_.Part } | Select-Object -Unique)) {
		$entry = $zip.GetEntry($name)
		if (-not $entry) { continue }
		$reader = New-Object IO.StreamReader($entry.Open())
		$text = $reader.ReadToEnd()
		$reader.Close()
		foreach ($edit in ($edits | Where-Object { $_.Part -eq $name })) { $text = [regex]::Replace($text, $edit.Find, $edit.With) }
		$entry.Delete()
		$writer = New-Object IO.StreamWriter($zip.CreateEntry($name).Open(), (New-Object Text.UTF8Encoding($false)))
		$writer.Write($text)
		$writer.Close()
	}
} finally { $zip.Dispose() }
Get-Item $out | Select-Object Name, Length
