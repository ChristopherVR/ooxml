# Manual native acceptance of shared style/color edits. Build inputs with the adjacent Bun script.
param([Parameter(Mandatory=$true)][string]$OutDir)
$ErrorActionPreference = 'Stop'
$manifest = Get-Content -Raw (Join-Path $OutDir 'manifest.json') | ConvertFrom-Json
$results = @()
$word = $null
$excel = $null
function Assert-Art($art, $expected) {
	$texts = @()
	for ($i = 1; $i -le $art.AllNodes.Count; $i++) { $texts += [string]$art.AllNodes.Item($i).TextFrame2.TextRange.Text }
	if (($texts -join '|') -ne ($expected.texts -join '|')) { throw "Native node text mismatch: $($texts -join '|')" }
	$color = [string]$art.Color.Id
	$style = [string]$art.QuickStyle.Id
	if ($color -ne $expected.colors) { throw "Native color mismatch: $color" }
	if ($style -ne $expected.style) { throw "Native style mismatch: $style" }
	return @{ texts = $texts; colors = $color; style = $style }
}
try {
	$word = New-Object -ComObject Word.Application
	$word.Visible = $false; $word.DisplayAlerts = 0
	foreach ($file in Get-ChildItem -LiteralPath $OutDir -Filter '*.docx') {
		$doc = $null
		try {
			$doc = $word.Documents.Open($file.FullName, $false, $true, $false)
			$result = Assert-Art $doc.InlineShapes.Item(1).SmartArt $manifest.($file.Name)
			$results += @{ file = $file.Name; version = [string]$word.Version; result = $result }
		} finally { if ($doc) { $doc.Close(0) } }
	}
	$word.Quit(); $word = $null
	$excel = New-Object -ComObject Excel.Application
	$excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.AskToUpdateLinks = $false
	foreach ($file in Get-ChildItem -LiteralPath $OutDir -Filter '*.xlsx') {
		$book = $null
		try {
			$before = @(Get-ChildItem ([IO.Path]::GetTempPath()) -Filter 'error*.xml' | ForEach-Object FullName)
			$book = $excel.Workbooks.Open($file.FullName, 0, $true)
			$newLogs = @(Get-ChildItem ([IO.Path]::GetTempPath()) -Filter 'error*.xml' | Where-Object { $before -notcontains $_.FullName })
			if ($newLogs.Count -gt 0) { throw "Excel produced a repair log: $($newLogs[0].FullName)" }
			if ([string]$excel.ActiveWindow.Caption -match 'Repaired') { throw 'Excel repaired the workbook.' }
			$shape = $book.Worksheets.Item(1).Shapes.Item(1)
			$result = Assert-Art $shape.SmartArt $manifest.($file.Name)
			$results += @{ file = $file.Name; version = [string]$excel.Version; result = $result }
		} finally { if ($book) { $book.Close($false) } }
	}
	$results | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $OutDir 'native-results.json')
	Write-Output "$($results.Count) native SmartArt formatting files passed."
} finally {
	if ($word) { $word.Quit() }
	if ($excel) { $excel.Quit() }
}
