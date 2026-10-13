param([Parameter(Mandatory = $true)][string]$OutputDirectory, [switch]$VerifyCore)
# Native Dynamic connector geometry evidence. Files containing Microsoft masters remain local.
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force $OutputDirectory | Out-Null
$out = (Resolve-Path $OutputDirectory).Path
$app = New-Object -ComObject Visio.InvisibleApp
$app.AlertResponse = 7
$results = @()
function CurveValues($shape) {
	$width = $shape.CellsU('Width').ResultIU; $height = $shape.CellsU('Height').ResultIU
	$values = @($shape.CellsU('Geometry1.E2').FormulaU.Trim() -replace '^NURBS\(|\)$', '' -split ',' | ForEach-Object { [double]::Parse($_.Trim(), [Globalization.CultureInfo]::InvariantCulture) })
	$result = @($shape.CellsU('Geometry1.X1').ResultIU, $shape.CellsU('Geometry1.Y1').ResultIU)
	for ($i = 4; $i -lt $values.Count; $i += 4) {
		$scaleX = if ($values[2] -eq 0) { $width } else { 1 }
		$scaleY = if ($values[3] -eq 0) { $height } else { 1 }
		$result += @(($values[$i] * $scaleX), ($values[$i + 1] * $scaleY))
	}
	$result += @($shape.CellsU('Geometry1.X2').ResultIU, $shape.CellsU('Geometry1.Y2').ResultIU,
		$shape.CellsU('Geometry1.A2').ResultIU, $values[0], $shape.CellsU('TxtPinX').ResultIU, $shape.CellsU('TxtPinY').ResultIU)
	return ,$result
}
try {
	if ($VerifyCore) {
		$expected = Get-Content (Join-Path $out 'curved-route-native.json') -Raw | ConvertFrom-Json
		$editedFiles = @(Get-ChildItem -LiteralPath $out -Filter 'edited-curve-*.vsdx')
		if ($editedFiles.Count -ne 10) { throw 'Generate all ten edited drawings before native acceptance.' }
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
			if ($reference.begin[0] -ne $reference.end[0] -and $reference.begin[1] -ne $reference.end[1]) {
				$native = $app.Documents.OpenEx((Join-Path $out "curve-$x-$y-1-1.vsdx"), 64)
				$nativeValues = CurveValues $native.Pages.Item(1).Shapes.ItemFromID(3)
				$editedValues = CurveValues $conn
				if ($nativeValues.Count -ne $editedValues.Count) { throw "Control count mismatch in $($file.Name)" }
				for ($i = 0; $i -lt $nativeValues.Count; $i++) {
					$tolerance = if ($i -ge $nativeValues.Count - 2) { 0.0002 } else { 1e-8 }
					if ([Math]::Abs($nativeValues[$i] - $editedValues[$i]) -gt $tolerance) { throw "Curve/text mismatch in $($file.Name), value $i`: $($editedValues[$i]) versus $($nativeValues[$i])" }
				}
				$native.Close()
			}
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
		$results += @{ target = $position; begin = @($conn.CellsU('BeginX').ResultIU, $conn.CellsU('BeginY').ResultIU); end = @($conn.CellsU('EndX').ResultIU, $conn.CellsU('EndY').ResultIU); text = @($conn.CellsU('TxtPinX').ResultIU, $conn.CellsU('TxtPinY').ResultIU); rows = $rows }
		$doc.SaveAs((Join-Path $out "curve-$x-$y-$width-$height.vsdx")) | Out-Null
		$doc.Close()
	}
	@{ version = $app.Version; cases = $results } | ConvertTo-Json -Depth 12 | Set-Content (Join-Path $out 'curved-route-native.json') -Encoding utf8
} finally {
	foreach ($document in @($app.Documents)) { $document.Saved = $true }
	$app.Quit()
}
