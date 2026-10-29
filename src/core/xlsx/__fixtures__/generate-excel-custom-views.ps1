# Generates excel-custom-views.xlsx with real Excel: two custom views ("Compact" with hidden rows
# and columns, a filter, 85% zoom, no gridlines and its own print setup; "Full" without hidden
# rows or print settings) on a two-sheet workbook. Run: pwsh -File generate-excel-custom-views.ps1
$ErrorActionPreference = 'Stop'
while (Get-Process EXCEL -ErrorAction SilentlyContinue) { Start-Sleep -Seconds 30 }
$out = Join-Path $PSScriptRoot 'excel-custom-views.xlsx'
$excel = New-Object -ComObject Excel.Application
$excel.DisplayAlerts = $false
try {
	$wb = $excel.Workbooks.Add()
	$ws = $wb.Worksheets.Item(1)
	$ws.Name = 'Data'
	$other = $wb.Worksheets.Add([Type]::Missing, $ws)
	$other.Name = 'Notes'
	$ws.Activate()
	foreach ($r in 1..6) { foreach ($c in 1..4) { $ws.Cells.Item($r, $c).Value2 = "R${r}C${c}" } }
	$excel.ActiveWindow.DisplayGridlines = $false
	$excel.ActiveWindow.Zoom = 85
	$ws.Rows.Item(3).Hidden = $true
	$ws.Columns.Item(3).Hidden = $true
	$ws.PageSetup.Orientation = 2
	$ws.PageSetup.CenterHeader = 'Compact'
	$ws.Range('A1:D6').AutoFilter(2, 'R2C2') | Out-Null
	[void]$wb.CustomViews.Add('Compact', $true, $true)
	$ws.AutoFilterMode = $false
	$ws.Rows.Item(3).Hidden = $false
	$ws.Columns.Item(3).Hidden = $false
	$excel.ActiveWindow.DisplayGridlines = $true
	$excel.ActiveWindow.Zoom = 100
	$ws.PageSetup.Orientation = 1
	$ws.PageSetup.CenterHeader = ''
	[void]$wb.CustomViews.Add('Full', $false, $false)
	if (Test-Path $out) { Remove-Item $out }
	$wb.SaveAs($out, 51)
	$wb.Close($false)
} finally {
	$excel.Quit()
	[void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}
