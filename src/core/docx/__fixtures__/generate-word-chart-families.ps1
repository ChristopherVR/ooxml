# Generates chart-families.docx: a combination chart (clustered columns plus a line on a secondary
# axis), a waterfall, a histogram, a bubble, an open-high-low-close stock, a treemap and a sunburst chart for the chart painter tests. Word cannot add these charts
# through its own object model reliably (extended charts fail outright), so each is authored in
# Excel and pasted into Word, which embeds the chart part, its workbook and the cached values the
# way Word saves them.
# Requires desktop Word and Excel (COM automation). Run from any folder:
#   pwsh -File src/core/docx/__fixtures__/generate-word-chart-families.ps1
$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$out = Join-Path $here 'chart-families.docx'

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
	$book = $excel.Workbooks.Add()
	$sheet = $book.Worksheets.Item(1)
	$sheet.Range('A1:C1').Value2 = [object[]]('Quarter', 'Revenue', 'Margin')
	$quarters = 'Q1', 'Q2', 'Q3', 'Q4'
	$revenue = 120, 150, 135, 190
	$margin = 0.12, 0.2, 0.15, 0.3
	for ($i = 0; $i -lt 4; $i++) {
		$sheet.Range("A$($i + 2)").Value2 = $quarters[$i]
		$sheet.Range("B$($i + 2)").Value2 = [double]$revenue[$i]
		$sheet.Range("C$($i + 2)").Value2 = [double]$margin[$i]
	}
	$sheet.Range('E1:F1').Value2 = [object[]]('Step', 'Amount')
	$steps = 'Start', 'Sales', 'Returns', 'Costs', 'Other', 'Total'
	$amounts = 100, 60, -20, -45, 15, 110
	for ($i = 0; $i -lt 6; $i++) {
		$sheet.Range("E$($i + 2)").Value2 = $steps[$i]
		$sheet.Range("F$($i + 2)").Value2 = [double]$amounts[$i]
	}
	$sheet.Range('H1').Value2 = 'Score'
	$scores = 3, 5, 7, 8, 8, 9, 12, 13, 13, 14, 15, 15, 15, 17, 18, 21, 22, 24, 27, 29
	for ($i = 0; $i -lt $scores.Count; $i++) { $sheet.Range("H$($i + 2)").Value2 = [double]$scores[$i] }

	# Bubble: x, y and size per point.
	$sheet.Range('J1:L1').Value2 = [object[]]('X', 'Y', 'Size')
	$bubbles = @(@(1, 3, 5), @(2, 5, 10), @(3, 4, 3), @(4, 8, 8))
	for ($i = 0; $i -lt 4; $i++) {
		for ($j = 0; $j -lt 3; $j++) { $sheet.Cells.Item($i + 2, 10 + $j).Value2 = [double]$bubbles[$i][$j] }
	}
	# Stock: open, high, low, close by day.
	$sheet.Range('N1:R1').Value2 = [object[]]('Day', 'Open', 'High', 'Low', 'Close')
	$days = 'Mon', 'Tue', 'Wed', 'Thu'
	$ohlc = @(@(10, 14, 9, 12), @(11, 15, 10, 13), @(12, 15, 10, 11), @(11, 14, 9, 13))
	for ($i = 0; $i -lt 4; $i++) {
		$sheet.Cells.Item($i + 2, 14).Value2 = $days[$i]
		for ($j = 0; $j -lt 4; $j++) { $sheet.Cells.Item($i + 2, 15 + $j).Value2 = [double]$ohlc[$i][$j] }
	}
	# Treemap and sunburst: a two-level hierarchy with a weight per leaf.
	$sheet.Range('T1:V1').Value2 = [object[]]('Group', 'Item', 'Weight')
	$produce = @(@('Fruit', 'Apple', 30), @('Fruit', 'Pear', 20), @('Veg', 'Leek', 10), @('Veg', 'Kale', 5), @('Nuts', 'Almond', 8))
	for ($i = 0; $i -lt 5; $i++) {
		$sheet.Cells.Item($i + 2, 20).Value2 = $produce[$i][0]
		$sheet.Cells.Item($i + 2, 21).Value2 = $produce[$i][1]
		$sheet.Cells.Item($i + 2, 22).Value2 = [double]$produce[$i][2]
	}

	# xlColumnClustered = 51, xlLineMarkers = 65, xlWaterfall = 119, xlHistogram = 118.
	$combo = $sheet.Shapes.AddChart2(-1, 51, 20, 200, 360, 220)
	$combo.Chart.SetSourceData($sheet.Range('A1:C5'))
	$combo.Chart.SeriesCollection(2).ChartType = 65
	$combo.Chart.SeriesCollection(2).AxisGroup = 2
	$combo.Chart.HasTitle = $true
	$combo.Chart.ChartTitle.Text = 'Revenue and margin'
	$waterfall = $sheet.Shapes.AddChart2(-1, 119, 400, 200, 360, 220)
	$waterfall.Chart.SetSourceData($sheet.Range('E1:F7'))
	$waterfall.Chart.HasTitle = $true
	$waterfall.Chart.ChartTitle.Text = 'Cash flow'
	$waterfall.Chart.SeriesCollection(1).Points(1).IsTotal = $true
	$waterfall.Chart.SeriesCollection(1).Points(6).IsTotal = $true
	$histogram = $sheet.Shapes.AddChart2(-1, 118, 780, 200, 360, 220)
	$histogram.Chart.SetSourceData($sheet.Range('H1:H21'))
	$histogram.Chart.HasTitle = $true
	$histogram.Chart.ChartTitle.Text = 'Scores'

	# xlBubble = 15, xlStockOHLC = 89, xlTreemap = 117, xlSunburst = 120.
	$bubble = $sheet.Shapes.AddChart2(-1, 15, 20, 440, 360, 220)
	# Explicit series: the header-less numeric columns would otherwise be read as three series.
	while ($bubble.Chart.SeriesCollection().Count -gt 0) { $bubble.Chart.SeriesCollection(1).Delete() | Out-Null }
	$points = $bubble.Chart.SeriesCollection().NewSeries()
	$points.Name = 'Points'
	$points.XValues = $sheet.Range('J2:J5')
	$points.Values = $sheet.Range('K2:K5')
	$points.BubbleSizes = '=' + $sheet.Name + '!$L$2:$L$5'
	$bubble.Chart.HasTitle = $true
	$bubble.Chart.ChartTitle.Text = 'Bubbles'
	$stock = $sheet.Shapes.AddChart2(-1, 51, 400, 440, 360, 220)
	$stock.Chart.SetSourceData($sheet.Range('N1:R5'))
	$stock.Chart.ChartType = 89
	$stock.Chart.HasTitle = $true
	$stock.Chart.ChartTitle.Text = 'Prices'
	$treemap = $sheet.Shapes.AddChart2(-1, 117, 780, 440, 360, 220)
	$treemap.Chart.SetSourceData($sheet.Range('T1:V6'))
	$treemap.Chart.HasTitle = $true
	$treemap.Chart.ChartTitle.Text = 'Produce treemap'
	$sunburst = $sheet.Shapes.AddChart2(-1, 120, 1160, 440, 360, 220)
	$sunburst.Chart.SetSourceData($sheet.Range('T1:V6'))
	$sunburst.Chart.HasTitle = $true
	$sunburst.Chart.ChartTitle.Text = 'Produce sunburst'

	$doc = $word.Documents.Add()
	$sel = $word.Selection
	$pieces = @(@('Combination', $combo), @('Waterfall', $waterfall), @('Histogram', $histogram), @('Bubble', $bubble), @('Stock', $stock), @('Treemap', $treemap), @('Sunburst', $sunburst))
	foreach ($piece in $pieces) {
		$sel.TypeText($piece[0])
		$sel.TypeParagraph()
		$piece[1].Copy()
		Start-Sleep -Seconds 1
		$sel.Paste()
		$sel.EndKey(6) | Out-Null
		$sel.TypeParagraph()
	}
	if (Test-Path $out) { Remove-Item $out }
	$doc.SaveAs2($out, 16)
	$doc.Close($false)
	$book.Close($false)
}
finally {
	$word.Quit()
	$excel.Quit()
	[void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word)
	[void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}
