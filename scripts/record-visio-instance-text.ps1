param([Parameter(Mandatory = $true)][string]$OutputDirectory)
# Records what native Visio saves when a stencil (master) instance gets text: before.vsdx holds a
# dropped Process and Decision glued by a Dynamic connector; after.vsdx is the same drawing once
# the Process has text and the Decision is wider. The captures embed Microsoft's stencil masters,
# so they are local evidence only and are never committed. Point the optional test at them with
# VISIO_NATIVE_INSTANCE_TEXT_DIR (src/core/visio/edit-text-instance.test.ts).
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force $OutputDirectory | Out-Null
$out = (Resolve-Path $OutputDirectory).Path
$app = New-Object -ComObject Visio.InvisibleApp
$app.AlertResponse = 7
try {
	$doc = $app.Documents.Add('')
	$stencil = $app.Documents.OpenEx('BASFLO_U.VSSX', 64)
	$page = $doc.Pages.Item(1)
	$process = $page.Drop($stencil.Masters.ItemU('Process'), 2, 6)
	$decision = $page.Drop($stencil.Masters.ItemU('Decision'), 2, 4)
	$process.AutoConnect($decision, 2)
	$doc.SaveAs((Join-Path $out 'before.vsdx')) | Out-Null
	$process.Text = 'Typed on an instance'
	$decision.CellsU('Width').FormulaU = '2 in'
	$doc.SaveAs((Join-Path $out 'after.vsdx')) | Out-Null
	"Recorded before.vsdx and after.vsdx in $out (Process is shape $($process.ID))."
} finally {
	foreach ($d in @($app.Documents)) { $d.Saved = $true }
	$app.Quit()
}
