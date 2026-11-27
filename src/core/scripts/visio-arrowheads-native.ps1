# Draws one 2 in line per BeginArrow/EndArrow code 1 to 45 in an invisible Visio, and exports
# the page as SVG (observable output only). src/core/visio/__fixtures__/arrowheads-native.md
# describes how the fixture is read from the export.
param([string]$Out)
$app = New-Object -ComObject Visio.InvisibleApp
try {
  $app.AlertResponse = 7
  $doc = $app.Documents.Add('')
  $page = $doc.Pages.Item(1)
  for ($code = 1; $code -le 45; $code++) {
    $col = [math]::Floor(($code - 1) / 15)
    $row = ($code - 1) % 15
    $x = 0.5 + $col * 2.7
    $y = 10.5 - $row * 0.7
    $line = $page.DrawLine($x, $y, $x + 2, $y)
    $line.CellsU('EndArrow').FormulaU = "$code"
    $line.CellsU('BeginArrow').FormulaU = "0"
    $line.CellsU('EndArrowSize').FormulaU = "2"
    $line.CellsU('LineWeight').FormulaU = "1 pt"
    $line.Text = "$code"
    $line.Name = "arrow$code"
  }
  $page.Export((Join-Path $Out 'arrows.svg'))
  $doc.SaveAs((Join-Path $Out 'arrows.vsdx')) | Out-Null
  $doc.Saved = $true
  $doc.Close()
} finally {
  $app.Quit()
}
