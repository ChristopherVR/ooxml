param(
	[Parameter(Mandatory = $true)][string]$OutputDirectory,
	[string]$Master = 'Angled connector'
)
# Records what native Visio writes for a straight line master from the Connectors stencil (not the
# Dynamic connector) glued to two flowchart shapes by their connection points, and what it writes
# after the second shape moves twice. 01.vsdx is the glued drawing; 02.vsdx and 03.vsdx follow the
# moves. The captures embed Microsoft masters, so they are local evidence only and never committed.
# The findings are asserted in src/core/visio/edit-stencil-line.test.ts. COM only: an invisible
# application, no keystrokes and no screen capture; it quits only the instance it starts.
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force $OutputDirectory | Out-Null
$out = (Resolve-Path $OutputDirectory).Path
$app = New-Object -ComObject Visio.InvisibleApp
$app.AlertResponse = 7
try {
	$doc = $app.Documents.Add('')
	$flow = $app.Documents.OpenEx('BASFLO_U.VSSX', 64 + 2)
	$connectors = $app.Documents.OpenEx('CONNEC_U.VSSX', 64 + 2)
	$process = $flow.Masters.ItemU('Process')
	$page = $doc.Pages.Item(1)
	$a = $page.Drop($process, 2, 6)
	$b = $page.Drop($process, 6, 4)
	$line = $page.Drop($connectors.Masters.ItemU($Master), 4, 5)
	# Connection point rows are visSectionConnectionPts (7); row 1 is the right side of Process.
	$line.CellsU('BeginX').GlueTo($a.CellsSRC(7, 1, 0))
	$line.CellsU('EndX').GlueTo($b.CellsSRC(7, 1, 0))
	$doc.SaveAs((Join-Path $out '01.vsdx')) | Out-Null
	$b.CellsU('PinX').ResultIU = 7
	$b.CellsU('PinY').ResultIU = 8
	$doc.SaveAs((Join-Path $out '02.vsdx')) | Out-Null
	$b.CellsU('PinX').ResultIU = 6
	$b.CellsU('PinY').ResultIU = 2
	$doc.SaveAs((Join-Path $out '03.vsdx')) | Out-Null
	"Recorded in $out."
} finally {
	foreach ($d in @($app.Documents)) { $d.Saved = $true }
	$app.Quit()
}
