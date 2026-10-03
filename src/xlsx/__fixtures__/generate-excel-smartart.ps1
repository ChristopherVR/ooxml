# Generates excel-smartart.xlsx (a SmartArt graphic authored by Excel) for the SmartArt tests.
# Requires desktop Excel (COM automation). Run from any folder:
#   pwsh -File src/xlsx/__fixtures__/generate-excel-smartart.ps1
# The account name is never changed; the author fields Excel stamps are replaced afterwards.
$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$out = Join-Path $here 'excel-smartart.xlsx'

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
	$wb = $excel.Workbooks.Add()
	$ws = $wb.Worksheets.Item(1)
	$ws.Range('A1').Value2 = 'SmartArt below'
	# Layout 1 is the Basic Block List.
	$shape = $ws.Shapes.AddSmartArt($excel.SmartArtLayouts.Item(1), 40, 30, 360, 220)
	$shape.Name = 'Process Diagram'
	$nodes = $shape.SmartArt.AllNodes
	while ($nodes.Count -gt 3) { $nodes.Item($nodes.Count).Delete() }
	$labels = @('Plan', 'Build', 'Ship')
	for ($i = 1; $i -le $nodes.Count; $i++) { $nodes.Item($i).TextFrame2.TextRange.Text = $labels[$i - 1] }
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
