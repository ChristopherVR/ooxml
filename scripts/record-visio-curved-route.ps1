param([Parameter(Mandatory = $true)][string]$OutputDirectory, [switch]$VerifyCore)
# Native Dynamic connector geometry evidence. Files containing Microsoft masters remain local.
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force $OutputDirectory | Out-Null
$out = (Resolve-Path $OutputDirectory).Path
$app = New-Object -ComObject Visio.InvisibleApp
$app.AlertResponse = 7
$results = @()
try {
	if ($VerifyCore) {
		$expected = Get-Content (Join-Path $out 'curved-route-native.json') -Raw | ConvertFrom-Json
		$editedFiles = @(Get-ChildItem -LiteralPath $out -Filter 'edited-curve-*.vsdx')
		if ($editedFiles.Count -ne 4) { throw 'Generate all four edited drawings before native acceptance.' }
		foreach ($file in $editedFiles) {
			$doc = $app.Documents.OpenEx($file.FullName, 64)
			$target = $doc.Pages.Item(1).Shapes.ItemFromID(2)
			$conn = $doc.Pages.Item(1).Shapes.ItemFromID(3)
			$x = $target.CellsU('PinX').ResultIU; $y = $target.CellsU('PinY').ResultIU
			$reference = $expected.cases | Where-Object { $_.target.Count -eq 2 -and $_.target[0] -eq $x -and $_.target[1] -eq $y } | Select-Object -First 1
			if (!$reference) { throw "Missing reference for $($file.Name)" }
			$actual = @($conn.CellsU('BeginX').ResultIU, $conn.CellsU('BeginY').ResultIU, $conn.CellsU('EndX').ResultIU, $conn.CellsU('EndY').ResultIU)
			$wanted = @($reference.begin) + @($reference.end)
			for ($i = 0; $i -lt 4; $i++) { if ([Math]::Abs($actual[$i] - $wanted[$i]) -gt 1e-8) { throw "Endpoint mismatch in $($file.Name): $actual versus $wanted" } }
			$doc.SaveAs((Join-Path $out "accepted-$($file.Name)")) | Out-Null
			$doc.Close()
			"Accepted $($file.Name): $actual"
		}
		return
	}
	foreach ($position in @(@(6, 3), @(6, 8), @(6, 6), @(6, 6.2), @(6, 6.8), @(6, 7), @(2, 3), @(2.2, 3), @(2.8, 3), @(3, 3), @(1, 3), @(0, 8), @(3.2, 6.2), @(6, 6.2, 1, 2), @(6, 6.8, 1, 2), @(2.2, 3, 2, 1), @(0, 6.2))) {
		$doc = $app.Documents.Add('')
		$page = $doc.Pages.Item(1)
		$a = $page.DrawRectangle(1.5, 5.5, 2.5, 6.5)
		$x = [double]$position[0]; $y = [double]$position[1]
		$width = 1; $height = 1
		if ($position.Count -eq 4) { $width = $position[2]; $height = $position[3] }
		$b = $page.DrawRectangle($x - $width / 2, $y - $height / 2, $x + $width / 2, $y + $height / 2)
		$a.AutoConnect($b, 0)
		$conn = $page.Shapes.Item($page.Shapes.Count)
		$conn.CellsU('ShapeRouteStyle').FormulaU = '1'
		$conn.CellsU('ConLineRouteExt').FormulaU = '2'
		$b.CellsU('PinY').ResultIU = $y + 0.0001
		$b.CellsU('PinY').ResultIU = $y
		$rows = @()
		for ($r = 1; $r -lt $conn.RowCount(10); $r++) {
			$cells = @()
			for ($c = 0; $c -lt 7; $c++) {
				try { $cells += @{ formula = $conn.CellsSRC(10, $r, $c).FormulaU; value = $conn.CellsSRC(10, $r, $c).ResultIU } } catch { $cells += $null }
			}
			$rows += @{ type = $conn.RowType(10, $r); cells = $cells }
		}
		$results += @{ target = $position; begin = @($conn.CellsU('BeginX').ResultIU, $conn.CellsU('BeginY').ResultIU); end = @($conn.CellsU('EndX').ResultIU, $conn.CellsU('EndY').ResultIU); rows = $rows }
		$doc.SaveAs((Join-Path $out "curve-$x-$y-$width-$height.vsdx")) | Out-Null
		$doc.Close()
	}
	@{ version = $app.Version; cases = $results } | ConvertTo-Json -Depth 12 | Set-Content (Join-Path $out 'curved-route-native.json') -Encoding utf8
} finally {
	foreach ($document in @($app.Documents)) { $document.Saved = $true }
	$app.Quit()
}
