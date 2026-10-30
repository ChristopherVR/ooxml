param([Parameter(Mandatory = $true)][string]$OutputDirectory)
# Records what native Visio saves when a control handle of a stencil instance is moved: Basic
# Shapes masters with a Controls section are dropped (before.vsdx), the free axis of their first
# control is set the way a drag writes it (after.vsdx) and manifest.json lists what was set. The
# files embed Microsoft's stencil masters, so they are local evidence only and are never
# committed. Point the optional test at them with VISIO_NATIVE_CONTROLS_DIR
# (src/core/visio/edit-control-native.test.ts). Visio runs invisibly; no keys are sent and
# nothing is captured.
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force $OutputDirectory | Out-Null
$out = (Resolve-Path $OutputDirectory).Path
$app = New-Object -ComObject Visio.InvisibleApp
$app.AlertResponse = 7
try {
	$doc = $app.Documents.Add('')
	$basic = $app.Documents.OpenEx('BASIC_U.VSSX', 64)
	$page = $doc.Pages.Item(1)
	$names = @('Can', 'Parallelogram', 'Trapezoid', 'Cross', 'Chevron', 'Cube', 'Drop', 'Cone')
	$dropped = @()
	$x = 1.5
	foreach ($name in $names) {
		$shape = $page.Drop($basic.Masters.ItemU($name), $x, 5)
		$x += 1.5
		if ($shape.CellExistsU('Controls.Row_1.Y', 0) -ne 0) { $dropped += , @($name, $shape) }
	}
	$doc.SaveAs((Join-Path $out 'before.vsdx')) | Out-Null
	$manifest = @()
	foreach ($item in $dropped) {
		$name = $item[0]; $shape = $item[1]
		$xFree = $shape.CellsU('Controls.Row_1.XCon').ResultIU -eq 0
		$yFree = $shape.CellsU('Controls.Row_1.YCon').ResultIU -eq 0
		$axis = if ($yFree) { 'Y' } elseif ($xFree) { 'X' } else { $null }
		if (-not $axis) { continue }
		$cell = $shape.CellsU("Controls.Row_1.$axis")
		$old = $cell.FormulaU
		$dimension = if ($axis -eq 'X') { 'Width' } else { 'Height' }
		$ratio = 0.3
		# The new position is a ratio of the shape size; Visio keeps the cell's own BOUND limits.
		$new = "$dimension*$ratio"
		$cell.FormulaU = $new
		$manifest += [ordered]@{ master = $name; shapeId = [string]$shape.ID; axis = $axis; ratio = $ratio; before = $old; after = $new; value = $cell.ResultIU }
	}
	$doc.SaveAs((Join-Path $out 'after.vsdx')) | Out-Null
	$manifest | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $out 'manifest.json') -Encoding utf8
	Write-Output ($manifest | ConvertTo-Json -Depth 5)
} finally {
	if ($doc) { $doc.Saved = $true }
	$app.Quit()
}
