param([Parameter(Mandatory = $true)][string]$OutputDirectory)
# Records what native Visio saves when a shape action runs: the Annotations "Callout - curved"
# master is dropped (action-before.vsdx) and its "Show Box" action (SETF(GetRef(Geometry2.NoShow),
# NOT(Geometry2.NoShow))) is triggered through the cell's Trigger method (action-after.vsdx). The
# files embed Microsoft's stencil masters, so they are local evidence only and are never
# committed. Point the optional test at them with VISIO_NATIVE_ACTIONS_DIR
# (src/core/visio/edit-action-native.test.ts). Visio runs invisibly; no keys are sent.
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force $OutputDirectory | Out-Null
$out = (Resolve-Path $OutputDirectory).Path
$app = New-Object -ComObject Visio.InvisibleApp
$app.AlertResponse = 7
try {
	$doc = $app.Documents.Add('')
	$annotations = $app.Documents.OpenEx('ANNOT_U.VSSX', 64)
	$page = $doc.Pages.Item(1)
	$shape = $page.Drop($annotations.Masters.ItemU('Callout - curved'), 4, 4)
	$doc.SaveAs((Join-Path $out 'action-before.vsdx')) | Out-Null
	$shape.CellsU('Actions.Row_1.Action').Trigger() | Out-Null
	$doc.SaveAs((Join-Path $out 'action-after.vsdx')) | Out-Null
	Write-Output ("menu after: " + $shape.CellsU('Actions.Row_1.Menu').ResultStr(''))
} finally {
	if ($doc) { $doc.Saved = $true }
	$app.Quit()
}
