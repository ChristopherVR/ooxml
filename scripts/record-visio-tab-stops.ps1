param([Parameter(Mandatory = $true)][string]$OutputDirectory)
# Records how native Visio 16 saves paragraph tab stops. Drops the Entity (ENTITY_M.VSSX) and Pivot
# Node (DRILLD_M.VSSX) masters, whose Tabs rows Microsoft saved with Position<n>/Alignment<n> cells,
# and saves the page (visio-tabs-sample.vsdx). Then writes what Visio did with a Tabs row added to a
# text shape: AddRow is refused for the Tabs section, so the row keeps its default cells. The
# masters are Microsoft's, so the files stay local evidence and are never committed. Visio runs
# invisibly; no keys are sent and nothing is captured.
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force $OutputDirectory | Out-Null
$out = (Resolve-Path $OutputDirectory).Path
$app = New-Object -ComObject Visio.InvisibleApp
$app.AlertResponse = 7
$doc = $null
try {
	$doc = $app.Documents.Add('')
	$page = $doc.Pages.Item(1)
	$x = 1
	foreach ($name in @('ENTITY_M.VSSX', 'DRILLD_M.VSSX')) {
		$stencil = $app.Documents.OpenEx($name, 64)
		$master = $stencil.Masters.ItemU(@{ 'ENTITY_M.VSSX' = 'Entity'; 'DRILLD_M.VSSX' = 'Pivot Node' }[$name])
		[void]$page.Drop($master, $x, 4)
		$x += 2
	}
	$doc.SaveAs((Join-Path $out 'visio-tabs-sample.vsdx')) | Out-Null

	$text = $page.DrawRectangle(1, 1, 3, 2)
	$text.Text = "a`tb`tc"
	try {
		[void]$text.AddRow(5, -1, 0)
		Write-Output 'AddRow on the Tabs section: accepted'
	} catch {
		Write-Output ('AddRow on the Tabs section: refused (' + $_.Exception.Message.Trim() + ')')
	}
	$text.CellsSRC(5, 0, 0).FormulaU = '1 in'
	Write-Output ('Tabs.X1 after a write: ' + $text.CellsSRC(5, 0, 0).FormulaU)
	$doc.SaveAs((Join-Path $out 'visio-tabs-text.vsdx')) | Out-Null
} finally {
	if ($doc) { $doc.Saved = $true }
	$app.Quit()
}
