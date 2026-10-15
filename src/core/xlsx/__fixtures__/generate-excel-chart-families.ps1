# Generates excel-chart-families.xlsx: a combination chart (clustered columns plus a line on a
# secondary axis), a waterfall and a histogram, authored by Excel for the chart painter tests.
# Requires desktop Excel (COM automation). Run from any folder:
#   pwsh -File src/core/xlsx/__fixtures__/generate-excel-chart-families.ps1
# Pass -Png <folder> to also export each chart as Excel draws it (reference images, not committed).
param([string]$Png = '')
$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$out = Join-Path $here 'excel-chart-families.xlsx'

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
	$wb = $excel.Workbooks.Add()
	$ws = $wb.Worksheets.Item(1)
	$ws.Name = 'Data'
	# Combo: revenue (columns) and margin (line, secondary axis) by quarter.
	$ws.Range('A1:C1').Value2 = [object[]]('Quarter', 'Revenue', 'Margin')
	$quarters = 'Q1', 'Q2', 'Q3', 'Q4'
	$revenue = 120, 150, 135, 190
	$margin = 0.12, 0.2, 0.15, 0.3
	for ($i = 0; $i -lt 4; $i++) {
		$ws.Range("A$($i + 2)").Value2 = $quarters[$i]
		$ws.Range("B$($i + 2)").Value2 = [double]$revenue[$i]
		$ws.Range("C$($i + 2)").Value2 = [double]$margin[$i]
	}
	# Waterfall: a start, changes and the closing total (marked as a subtotal below).
	$ws.Range('E1:F1').Value2 = [object[]]('Step', 'Amount')
	$steps = 'Start', 'Sales', 'Returns', 'Costs', 'Other', 'Total'
	$amounts = 100, 60, -20, -45, 15, 110
	for ($i = 0; $i -lt 6; $i++) {
		$ws.Range("E$($i + 2)").Value2 = $steps[$i]
		$ws.Range("F$($i + 2)").Value2 = [double]$amounts[$i]
	}
	# Histogram: raw observations.
	$ws.Range('H1').Value2 = 'Score'
	$scores = 3, 5, 7, 8, 8, 9, 12, 13, 13, 14, 15, 15, 15, 17, 18, 21, 22, 24, 27, 29
	for ($i = 0; $i -lt $scores.Count; $i++) { $ws.Range("H$($i + 2)").Value2 = [double]$scores[$i] }

	$combo = $ws.Shapes.AddChart2(-1, 51, 20, 200, 360, 220)
	$combo.Chart.SetSourceData($ws.Range('A1:C5'))
	$combo.Chart.SeriesCollection(2).ChartType = 65
	$combo.Chart.SeriesCollection(2).AxisGroup = 2
	$combo.Chart.HasTitle = $true
	$combo.Chart.ChartTitle.Text = 'Revenue and margin'

	$waterfall = $ws.Shapes.AddChart2(-1, 119, 400, 200, 360, 220)
	$waterfall.Chart.SetSourceData($ws.Range('E1:F7'))
	$waterfall.Chart.HasTitle = $true
	$waterfall.Chart.ChartTitle.Text = 'Cash flow'
	$waterfall.Chart.SeriesCollection(1).Points(6).IsTotal = $true
	$waterfall.Chart.SeriesCollection(1).Points(1).IsTotal = $true

	$histogram = $ws.Shapes.AddChart2(-1, 118, 780, 200, 360, 220)
	$histogram.Chart.SetSourceData($ws.Range('H1:H21'))
	$histogram.Chart.HasTitle = $true
	$histogram.Chart.ChartTitle.Text = 'Scores'

	if ($Png) {
		New-Item -ItemType Directory -Force $Png | Out-Null
		$combo.Chart.Export((Join-Path $Png 'combo.png'))
		$waterfall.Chart.Export((Join-Path $Png 'waterfall.png'))
		$histogram.Chart.Export((Join-Path $Png 'histogram.png'))
	}
	if (Test-Path $out) { Remove-Item $out }
	$wb.SaveAs($out, 51)
	$wb.Close($false)
}
finally {
	$excel.Quit()
	[void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}
